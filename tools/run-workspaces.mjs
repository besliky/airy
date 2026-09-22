import { spawn } from 'node:child_process'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { cpus } from 'node:os'
import { dirname, join } from 'node:path'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'

const repoRoot = dirname(dirname(fileURLToPath(import.meta.url)))

// Canonical workspace order, carried over verbatim from the serial `test` and
// `typecheck` chains that used to live in the root package.json scripts
// (packages first, then apps, in dependency-ish order). The self-check below
// fails the run when a workspace with the requested script exists outside this
// list or a listed workspace lacks it, so a new workspace cannot be silently
// dropped from `npm test` / `npm run typecheck`.
const WORKSPACE_ORDER = [
  'i18n',
  'electron-utils',
  'font-metrics',
  'docx-engine',
  'pdf2docx',
  'html2docx',
  'file-parse',
  'pptx-engine',
  'pptx-render',
  'ai-search',
  'agent-core',
  'ai-provider',
  'project-store',
  'ui',
  'mcp-server',
  'docs',
  'sheets',
  'shell',
  'slides',
  'pdf',
  'markdown',
  'html',
]

// CI shard coverage guard: `--require-all --excluded docs sheets` fails when a
// workspace is added to WORKSPACE_ORDER but not to any CI shard (the same
// silent-drop failure mode the self-check below guards against).
const CI_EXCLUDED_FROM_SHARDS = ['docs', 'sheets']

// Workspace build order, carried over verbatim from the `build:all` chain in
// the root package.json. The builds are independent (apps only reference each
// other's output at runtime), so the CI e2e job runs this list in parallel
// instead of the serial chain. Locally `npm run build:all` keeps the serial
// chain unchanged.
const BUILD_WORKSPACE_ORDER = [
  'mcp-server',
  'docs',
  'sheets',
  'slides',
  'pdf',
  'markdown',
  'html',
  'shell',
]

// Workspaces that must not run at the same time as each other because they
// contend on a shared resource (fixed ports, shared output dirs, ...). Members
// of a group run sequentially (in WORKSPACE_ORDER) while unrelated workspaces
// keep running in parallel. Add a group here — not a full `--concurrency 1` —
// when two suites are found to collide. (The current suites had no such
// collisions: their only parallel-run friction is CPU starvation, which the
// halved default concurrency below absorbs.)
const CONFLICT_GROUPS = []

const COLORS = ['36', '33', '35', '32', '34', '31']
const useColor = process.stdout.isTTY && !process.env.NO_COLOR

function tag(name, index) {
  const color = COLORS[index % COLORS.length]
  return useColor ? `\x1b[${color}m[${name}]\x1b[0m` : `[${name}]`
}

function usage() {
  console.error(
    'Usage: node tools/run-workspaces.mjs <test|typecheck|build> [flags]\n' +
      'Flags: --concurrency <n> | --only <names...> | --excluded <names...> | ' +
      '--require-all | --dry-run. Flags may come before or after the mode. ' +
      'Concurrency defaults to min(workspace count, max(2, cpu count / 2)); ' +
      'AIRY_WORKSPACE_CONCURRENCY overrides the default. --only restricts the ' +
      'run to the named workspaces (CI shards); --require-all asserts that ' +
      '--only plus --excluded covers the whole order (shard coverage guard); ' +
      '--dry-run validates and prints the selection without running it.',
  )
}

// Order-independent parsing: flags (and their values) are lifted out wherever
// they appear, so `--concurrency 4 test` no longer misreads `4` as the mode.
const args = process.argv.slice(2)
let concurrency = Number.parseInt(process.env.AIRY_WORKSPACE_CONCURRENCY ?? '', 10)
const positional = []

function takeNames(startIndex) {
  const names = []
  let index = startIndex
  while (index < args.length && !args[index].startsWith('-')) {
    // Accept both space- and comma-separated lists (`--only a b`, `--only a,b`).
    for (const name of args[index].split(',')) {
      if (name !== '') names.push(name)
    }
    index += 1
  }
  return [names, index]
}

let onlyNames = []
let excludedNames = []
let requireAll = false
let dryRun = false
for (let i = 0; i < args.length; i++) {
  const arg = args[i]
  if (arg === '--concurrency' || arg === '-c') {
    concurrency = Number.parseInt(args[++i], 10)
  } else if (arg.startsWith('--concurrency=')) {
    concurrency = Number.parseInt(arg.slice(arg.indexOf('=') + 1), 10)
  } else if (arg === '--only' || arg === '--excluded') {
    const flag = arg
    const [names, next] = takeNames(i + 1)
    if (names.length === 0) {
      console.error(`${flag} requires at least one workspace name`)
      process.exit(2)
    }
    if (flag === '--only') onlyNames = names
    else excludedNames = names
    i = next - 1
  } else if (arg === '--require-all') {
    requireAll = true
  } else if (arg === '--dry-run') {
    dryRun = true
  } else if (arg.startsWith('-')) {
    continue
  } else {
    positional.push(arg)
  }
}
if (requireAll && onlyNames.length === 0) {
  console.error('--require-all is only meaningful together with --only')
  process.exit(2)
}
const mode = positional[0]
if (!Number.isFinite(concurrency) || concurrency < 1) concurrency = undefined
if (mode !== 'test' && mode !== 'typecheck' && mode !== 'build') {
  usage()
  process.exit(2)
}
const order = mode === 'build' ? BUILD_WORKSPACE_ORDER : WORKSPACE_ORDER

// --- Self-check: the list must exactly cover the workspaces that have the
// script, so nothing is silently dropped (the failure mode of hand-maintained
// `&&` chains). The mcp-server directory publishes as @airy-office/mcp.
function workspaceDir(name) {
  for (const parent of ['packages', 'apps']) {
    const dir = join(repoRoot, parent, name)
    const manifest = join(dir, 'package.json')
    if (existsSync(manifest)) return { dir, manifest }
  }
  return null
}

function discoverWorkspacesWithScript(script) {
  const found = []
  for (const parent of ['packages', 'apps']) {
    for (const entry of readdirSync(join(repoRoot, parent), { withFileTypes: true })) {
      if (!entry.isDirectory()) continue
      const manifest = join(repoRoot, parent, entry.name, 'package.json')
      if (!existsSync(manifest)) continue
      const pkg = JSON.parse(readFileSync(manifest, 'utf8'))
      if (pkg.scripts?.[script]) found.push(entry.name)
    }
  }
  return found
}

const problems = []
const discovered = discoverWorkspacesWithScript(mode)
if (onlyNames.length > 0) {
  // Shard mode: validate the requested subset instead of the full order.
  const unique = [...new Set(onlyNames)]
  const unknown = unique.filter((name) => !order.includes(name))
  for (const name of unknown) problems.push(`workspace "${name}" is not in the ${mode} order`)
  for (const name of unique) {
    const location = workspaceDir(name)
    if (!location) problems.push(`workspace "${name}" does not exist`)
    else if (!JSON.parse(readFileSync(location.manifest, 'utf8')).scripts?.[mode])
      problems.push(`workspace "${name}" has no "${mode}" script`)
  }
  if (requireAll) {
    // Coverage guard: every workspace must run in some CI shard.
    const covered = new Set([...unique, ...CI_EXCLUDED_FROM_SHARDS, ...excludedNames])
    for (const name of order) {
      if (!covered.has(name))
        problems.push(`workspace "${name}" is not covered by any CI test shard`)
    }
    for (const name of excludedNames) {
      if (!order.includes(name)) problems.push(`workspace "${name}" is not in the ${mode} order`)
    }
  }
  if (problems.length === 0 && dryRun) {
    console.log(
      `dry-run ${mode}: would run ${unique.length} workspaces (ordered): ${unique
        .slice()
        .sort((a, b) => order.indexOf(a) - order.indexOf(b))
        .join(', ')}`,
    )
    process.exit(0)
  }
} else {
  for (const name of order) {
    const location = workspaceDir(name)
    if (!location) problems.push(`workspace "${name}" from the ${mode} order does not exist`)
    else if (!JSON.parse(readFileSync(location.manifest, 'utf8')).scripts?.[mode])
      problems.push(`workspace "${name}" has no "${mode}" script`)
  }
  for (const name of discovered) {
    if (!order.includes(name))
      problems.push(
        `workspace "${name}" has a "${mode}" script but is missing from the ${mode} order in tools/run-workspaces.mjs — add it or it will never run`,
      )
  }
  if (problems.length === 0 && dryRun) {
    console.log(`dry-run ${mode}: would run all ${order.length} workspaces`)
    process.exit(0)
  }
}
if (problems.length > 0) {
  console.error(`run-workspaces self-check failed for "${mode}":`)
  for (const problem of problems) console.error(`  - ${problem}`)
  process.exit(2)
}

// Default concurrency: each workspace's vitest run already parallelizes test
// files across the whole machine, so running one workspace per CPU
// oversubscribes the box 3-4x and starves load-sensitive tests (20s vitest
// timeouts on CPU-heavy cases, wasm conversions). Half a workspace per CPU
// kept the full suite green on a 12-core dev box while still cutting wall
// time ~30%; override with --concurrency / AIRY_WORKSPACE_CONCURRENCY.
const limit = concurrency ?? Math.min(order.length, Math.max(2, Math.floor(cpus().length / 2)))
const effectiveLimit = Math.min(limit, order.length)

// --- Scheduler: start tasks in order, never exceeding the concurrency limit,
// and hold back any task whose conflict-group members are still running.
const conflictGroupsOf = (name) =>
  CONFLICT_GROUPS.filter((group) => group.includes(name)).map((group) => new Set(group))

// Shard mode runs only the requested subset, in canonical order.
const scheduled =
  onlyNames.length > 0
    ? [...new Set(onlyNames)]
        .slice()
        .sort((a, b) => order.indexOf(a) - order.indexOf(b))
        .map((name) => ({ name, groups: conflictGroupsOf(name) }))
    : order.map((name) => ({ name, groups: conflictGroupsOf(name) }))
// Constant: `scheduled` is spliced as tasks start, so finish() must not
// compare against its live length (that fires mid-run once completed and
// pending counts cross).
const taskCount = scheduled.length
const running = []
const results = []

function wouldConflict(task) {
  return running.some((other) => task.groups.some((group) => group.has(other.name)))
}

function pump() {
  while (running.length < effectiveLimit) {
    const next = scheduled.find((task) => !wouldConflict(task))
    if (!next) return
    scheduled.splice(scheduled.indexOf(next), 1)
    start(next)
  }
}

function start(task) {
  const index = order.indexOf(task.name)
  const label = tag(task.name, index)
  const startedAt = Date.now()
  const child = spawn(
    process.platform === 'win32' ? 'npm.cmd' : 'npm',
    ['run', mode, '-w', `@airy-office/${task.name === 'mcp-server' ? 'mcp' : task.name}`],
    { cwd: repoRoot, shell: process.platform === 'win32' },
  )
  const entry = { name: task.name, child }
  running.push(entry)
  console.log(`${label} start (${running.length}/${effectiveLimit})`)
  const pipe = (stream, out) => {
    createInterface({ input: stream }).on('line', (line) => out.write(`${label} ${line}\n`))
  }
  pipe(child.stdout, process.stdout)
  pipe(child.stderr, process.stderr)
  let settled = false
  const settle = (code) => {
    if (settled) return
    settled = true
    const seconds = ((Date.now() - startedAt) / 1000).toFixed(1)
    running.splice(running.indexOf(entry), 1)
    const ok = code === 0
    results.push({ name: task.name, ok, seconds })
    console.log(`${label} ${ok ? 'done' : 'FAILED'} in ${seconds}s`)
    if (results.length === taskCount) finish()
    else pump()
  }
  child.on('exit', (code) => settle(code))
  child.on('error', (error) => {
    console.error(`${label} could not run: ${error.message}`)
    settle(1)
  })
}

function finish() {
  const failed = results.filter((r) => !r.ok)
  console.log(
    `\n${mode}: ${results.length - failed.length}/${results.length} workspaces passed in ${((Date.now() - startedTotalAt) / 1000).toFixed(1)}s`,
  )
  if (failed.length > 0) {
    console.error(`failed workspaces: ${failed.map((r) => r.name).join(', ')}`)
    process.exitCode = 1
  }
}

const startedTotalAt = Date.now()
pump()

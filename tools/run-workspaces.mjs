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
    'Usage: node tools/run-workspaces.mjs <test|typecheck> [--concurrency <n>]\n' +
      'Flags may come before or after the mode. Concurrency defaults to ' +
      'min(workspace count, max(2, cpu count / 2)); ' +
      'AIRY_WORKSPACE_CONCURRENCY overrides the default.',
  )
}

// Order-independent parsing: flags (and their values) are lifted out wherever
// they appear, so `--concurrency 4 test` no longer misreads `4` as the mode.
const args = process.argv.slice(2)
let concurrency = Number.parseInt(process.env.AIRY_WORKSPACE_CONCURRENCY ?? '', 10)
const positional = []
for (let i = 0; i < args.length; i++) {
  const arg = args[i]
  if (arg === '--concurrency' || arg === '-c') {
    concurrency = Number.parseInt(args[++i], 10)
  } else if (arg.startsWith('--concurrency=')) {
    concurrency = Number.parseInt(arg.slice(arg.indexOf('=') + 1), 10)
  } else if (arg.startsWith('-')) {
    continue
  } else {
    positional.push(arg)
  }
}
const mode = positional[0]
if (!Number.isFinite(concurrency) || concurrency < 1) concurrency = undefined
if (mode !== 'test' && mode !== 'typecheck') {
  usage()
  process.exit(2)
}

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
for (const name of WORKSPACE_ORDER) {
  const location = workspaceDir(name)
  if (!location) problems.push(`workspace "${name}" from WORKSPACE_ORDER does not exist`)
  else if (!JSON.parse(readFileSync(location.manifest, 'utf8')).scripts?.[mode])
    problems.push(`workspace "${name}" has no "${mode}" script`)
}
for (const name of discovered) {
  if (!WORKSPACE_ORDER.includes(name))
    problems.push(
      `workspace "${name}" has a "${mode}" script but is missing from WORKSPACE_ORDER in tools/run-workspaces.mjs — add it or its tests will never run`,
    )
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
const limit =
  concurrency ?? Math.min(WORKSPACE_ORDER.length, Math.max(2, Math.floor(cpus().length / 2)))
const effectiveLimit = Math.min(limit, WORKSPACE_ORDER.length)

// --- Scheduler: start tasks in order, never exceeding the concurrency limit,
// and hold back any task whose conflict-group members are still running.
const conflictGroupsOf = (name) =>
  CONFLICT_GROUPS.filter((group) => group.includes(name)).map((group) => new Set(group))

const pending = WORKSPACE_ORDER.map((name) => ({ name, groups: conflictGroupsOf(name) }))
const running = []
const results = []

function wouldConflict(task) {
  return running.some((other) => task.groups.some((group) => group.has(other.name)))
}

function pump() {
  while (running.length < effectiveLimit) {
    const next = pending.find((task) => !wouldConflict(task))
    if (!next) return
    pending.splice(pending.indexOf(next), 1)
    start(next)
  }
}

function start(task) {
  const index = WORKSPACE_ORDER.indexOf(task.name)
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
    if (results.length === WORKSPACE_ORDER.length) finish()
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

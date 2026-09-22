// CI matrix planner for the affected-only PR strategy (PERF-1550 phase 2).
//
// Emits the `test` and `e2e` job matrices as GitHub Actions outputs, consumed
// by .github/workflows/ci.yml via `matrix: ${{ fromJson(needs.plan.outputs.*) }}`:
//   test — vitest legs: the engines/apps workspace groups plus the docs/sheets
//          vitest shards (with their Rust sidecar rebuild)
//   e2e  — the Electron Playwright shards
//
// pull_request runs get the affected subset only; push and workflow_dispatch
// (and any diff we cannot confidently narrow) always get the full matrices, so
// the dispatch path and main-branch runs are never weakened.
//
// Rules:
//   - any changed path outside apps/** and packages/** (root configs, .github,
//     tools, scripts, fixtures, e2e specs, ...) → full matrices: outside the
//     workspaces we cannot attribute the change, so we fail wide;
//   - a workspace is affected when changed itself, or depended on
//     (transitively) by a changed workspace;
//   - e2e shards run iff one of the app workspaces (workspaces under apps/
//     with a build script) is affected. Every package in the repo is pulled
//     into some app's bundle except the standalone test targets
//     (html2docx/pdf2docx/mcp-server), which is exactly the set of PRs that
//     can skip the ~4m Electron suite.
//
// The workspace graph is read live from the package.json manifests
// (dependencies + devDependencies, @airy-office/* scopes only) — the same
// discovery tools/run-workspaces.mjs uses — so the graph cannot drift from
// the repo.
//
// The full test matrix below is the single source of truth for the shard
// coverage guard: it is validated against the canonical workspace order on
// every run (run-workspaces test --dry-run --require-all), so a workspace
// missing from a group fails the plan instead of silently never running.
//
// Usage:
//   node tools/ci-plan.mjs                     # in CI: event from GITHUB_EVENT_NAME
//   node tools/ci-plan.mjs --event pull_request --base HEAD^1
//   node tools/ci-plan.mjs --event pull_request --paths /tmp/changed-paths.txt
// The --paths file lists one changed path per line and replaces the git diff
// (local dry-runs).

import { execFileSync } from 'node:child_process'
import { appendFileSync, existsSync, readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = dirname(dirname(fileURLToPath(import.meta.url)))

// Unit-test workspace groups of the CI test matrix; docs and sheets are heavy
// enough to get their own vitest shards (mirrored in the workflow). Keep in
// sync with .github/workflows/ci.yml — the coverage guard below fails the run
// when this list and the real workspace set drift apart.
const ENGINES_GROUP = [
  'html2docx',
  'docx-engine',
  'markdown',
  'html',
  'pptx-render',
  'i18n',
  'electron-utils',
  'font-metrics',
  'ai-search',
  'agent-core',
  'ai-provider',
  'project-store',
  'ui',
]
const APPS_TEST_GROUP = [
  'slides',
  'pptx-engine',
  'pdf',
  'mcp-server',
  'pdf2docx',
  'file-parse',
  'shell',
]

// Number of e2e (Playwright shard) legs in the full matrix. The workflow name
// template and bundle-gate condition key off matrix.shard == 1 / job-total.
const E2E_SHARDS = 3

// --- Order-insensitive flags, same style as tools/run-workspaces.mjs.
const args = process.argv.slice(2)
let flagEvent = ''
let flagBase = ''
let flagPaths = ''
for (let i = 0; i < args.length; i++) {
  const arg = args[i]
  if (arg === '--event') flagEvent = args[++i] ?? ''
  else if (arg.startsWith('--event=')) flagEvent = arg.slice(arg.indexOf('=') + 1)
  else if (arg === '--base') flagBase = args[++i] ?? ''
  else if (arg.startsWith('--base=')) flagBase = arg.slice(arg.indexOf('=') + 1)
  else if (arg === '--paths') flagPaths = args[++i] ?? ''
  else if (arg.startsWith('--paths=')) flagPaths = arg.slice(arg.indexOf('=') + 1)
  else {
    console.error(`unexpected argument: ${arg}`)
    process.exit(2)
  }
}
const event = flagEvent || process.env.GITHUB_EVENT_NAME || 'push'
if (!['push', 'pull_request', 'workflow_dispatch'].includes(event)) {
  console.error(`unknown event: ${event}`)
  process.exit(2)
}

// --- Workspace discovery: live from the manifests, so new workspaces and
// dependency edges are picked up without editing this file.
function discoverWorkspaces() {
  const nameToDir = new Map()
  const manifests = []
  for (const parent of ['packages', 'apps']) {
    for (const entry of readdirSync(join(repoRoot, parent), { withFileTypes: true })) {
      if (!entry.isDirectory()) continue
      const manifest = join(repoRoot, parent, entry.name, 'package.json')
      if (!existsSync(manifest)) continue
      const pkg = JSON.parse(readFileSync(manifest, 'utf8'))
      if (pkg.name) nameToDir.set(pkg.name, entry.name)
      manifests.push({ dir: entry.name, parent, pkg })
    }
  }
  const workspaces = new Map()
  for (const { dir, parent, pkg } of manifests) {
    const deps = [
      ...new Set(
        Object.keys({ ...pkg.dependencies, ...pkg.devDependencies })
          .filter((name) => name.startsWith('@airy-office/'))
          .map((name) => nameToDir.get(name))
          .filter((target) => target && target !== dir),
      ),
    ]
    // App = the Electron-suite-relevant unit: an apps/ workspace with a build
    // script (docs, html, markdown, pdf, sheets, shell, slides). mcp-server
    // has a build script but lives under packages/ — it is not loaded by the
    // shell and must not gate e2e.
    workspaces.set(dir, { deps, app: parent === 'apps' && Boolean(pkg.scripts?.build) })
  }
  return workspaces
}

const workspaces = discoverWorkspaces()

for (const group of [ENGINES_GROUP, APPS_TEST_GROUP]) {
  for (const name of group) {
    if (!workspaces.has(name)) {
      console.error(`ci-plan workspace group lists "${name}", which does not exist in the repo`)
      process.exit(2)
    }
  }
}

// Shard coverage guard: the groups above (plus docs/sheets, which shard their
// vitest suites) must cover the canonical workspace order — a workspace with a
// test script missing from every group fails here instead of silently never
// running in CI. Reuses the exact guard command the quality job ran in phase 1.
execFileSync(
  process.execPath,
  [
    'tools/run-workspaces.mjs',
    'test',
    '--dry-run',
    '--require-all',
    '--excluded',
    'docs',
    'sheets',
    '--only',
    ...ENGINES_GROUP,
    ...APPS_TEST_GROUP,
  ],
  { cwd: repoRoot, stdio: 'inherit' },
)

// --- Changed paths: git diff against the PR base (the plan checkout is the
// merge commit, so merge-base resolves to its first parent — the exact main
// revision the PR was built against).
function changedPathsFromGit(base) {
  try {
    const mergeBase = execFileSync('git', ['merge-base', base, 'HEAD'], {
      cwd: repoRoot,
      encoding: 'utf8',
    }).trim()
    return execFileSync('git', ['diff', '--name-only', mergeBase, 'HEAD'], {
      cwd: repoRoot,
      encoding: 'utf8',
    })
      .split('\n')
      .filter((line) => line !== '')
  } catch {
    return null
  }
}

function classifyPaths(paths) {
  const changed = new Set()
  let e2eTrigger = false
  for (const path of paths) {
    const segments = path.split('/')
    if ((segments[0] === 'apps' || segments[0] === 'packages') && segments[1]) {
      if (!workspaces.has(segments[1])) return { full: `unrecognized workspace path: ${path}` }
      changed.add(segments[1])
    } else if (segments[0] === 'e2e') {
      e2eTrigger = true
    } else {
      return { full: `path outside workspaces: ${path}` }
    }
  }
  return { changed, e2eTrigger }
}

// Affected closure: everything that (transitively) depends on a changed
// workspace. A rebuild of docs/sheets/renderer bundles is driven by its
// dependencies, so dependents must re-run, not dependencies.
function affectedClosure(changed) {
  const dependents = new Map()
  for (const [dir, { deps }] of workspaces) {
    for (const dep of deps) {
      if (!dependents.has(dep)) dependents.set(dep, new Set())
      dependents.get(dep).add(dir)
    }
  }
  const affected = new Set()
  const queue = [...changed]
  while (queue.length > 0) {
    const current = queue.pop()
    if (affected.has(current)) continue
    affected.add(current)
    queue.push(...(dependents.get(current) ?? []))
  }
  return affected
}

// --- Matrices. The full one must stay identical to the phase-1 static matrix.
function fullTestMatrix() {
  return {
    include: [
      { name: 'engines', workspaces: ENGINES_GROUP.join(' ') },
      { name: 'apps', workspaces: APPS_TEST_GROUP.join(' ') },
      { name: 'docs-shard-1', workspaces: 'docs', vitestShard: '1/2' },
      { name: 'docs-shard-2', workspaces: 'docs', vitestShard: '2/2' },
      { name: 'sheets-shard-1', workspaces: 'sheets', vitestShard: '1/2', rust: true },
      { name: 'sheets-shard-2', workspaces: 'sheets', vitestShard: '2/2', rust: true },
    ],
  }
}

function affectedTestMatrix(affected) {
  const legs = []
  const engines = ENGINES_GROUP.filter((name) => affected.has(name))
  if (engines.length > 0) legs.push({ name: 'engines', workspaces: engines.join(' ') })
  const apps = APPS_TEST_GROUP.filter((name) => affected.has(name))
  if (apps.length > 0) legs.push({ name: 'apps', workspaces: apps.join(' ') })
  if (affected.has('docs'))
    legs.push(
      { name: 'docs-shard-1', workspaces: 'docs', vitestShard: '1/2' },
      { name: 'docs-shard-2', workspaces: 'docs', vitestShard: '2/2' },
    )
  if (affected.has('sheets'))
    legs.push(
      { name: 'sheets-shard-1', workspaces: 'sheets', vitestShard: '1/2', rust: true },
      { name: 'sheets-shard-2', workspaces: 'sheets', vitestShard: '2/2', rust: true },
    )
  return { include: legs }
}

function e2eMatrix(run) {
  const include = []
  if (run) for (let shard = 1; shard <= E2E_SHARDS; shard++) include.push({ shard })
  return { include }
}

// --- Decide.
let test
let e2e
let mode
if (event !== 'pull_request') {
  mode = `full (${event} always runs everything)`
  test = fullTestMatrix()
  e2e = e2eMatrix(true)
} else {
  const paths = flagPaths
    ? readFileSync(flagPaths, 'utf8')
        .split('\n')
        .map((line) => line.trim())
        .filter((line) => line !== '')
    : changedPathsFromGit(flagBase || 'HEAD^1')
  if (paths === null) {
    mode = `full (could not diff against base "${flagBase || 'HEAD^1'}")`
    test = fullTestMatrix()
    e2e = e2eMatrix(true)
  } else if (paths.length === 0) {
    mode = 'full (empty diff)'
    test = fullTestMatrix()
    e2e = e2eMatrix(true)
  } else {
    const classified = classifyPaths(paths)
    if (classified.full) {
      mode = `full (${classified.full})`
      test = fullTestMatrix()
      e2e = e2eMatrix(true)
    } else {
      const affected = affectedClosure(classified.changed)
      const affectedApps = [...affected].filter((name) => workspaces.get(name).app)
      mode = `affected (${affected.size} workspaces: ${[...affected].sort().join(', ')})`
      test = affectedTestMatrix(affected)
      e2e = e2eMatrix(classified.e2eTrigger || affectedApps.length > 0)
      if (affectedApps.length > 0) mode += `; apps: ${affectedApps.sort().join(', ')}`
      else if (e2e.include.length > 0) mode += '; apps: none (e2e/ changed)'
    }
  }
}

// --- Emit. GitHub skips a fromJson matrix job whose include array is empty,
// which is exactly the wanted behavior for e2e-less PRs.
const outputs = { test: JSON.stringify(test), e2e: JSON.stringify(e2e) }
console.log(`[ci-plan] event=${event} → ${mode}`)
console.log(
  `[ci-plan] test matrix (${test.include.length} legs): ` +
    test.include
      .map((leg) => `${leg.name}${leg.vitestShard ? ` [shard ${leg.vitestShard}]` : ''}`)
      .join(', '),
)
console.log(`[ci-plan] e2e matrix: ${e2e.include.length} shards`)
if (process.env.GITHUB_OUTPUT) {
  appendFileSync(process.env.GITHUB_OUTPUT, `test=${outputs.test}\ne2e=${outputs.e2e}\n`)
}

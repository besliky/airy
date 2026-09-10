/**
 * Egress guard: no functional Genspark leftovers may return to the fork's
 * source trees. Phase 4a removed the account login, the genspark AI
 * provider, the gsk CLI backend, the auto-updater and usage analytics;
 * this gate fails the build when any of the network endpoints or
 * dependencies they used reappear:
 *
 *   - genspark.ai / genspark.com URL domains (any subdomain, any path)
 *   - the retired login endpoints (device-code / api_tokens /
 *     office_addin_auth path segments)
 *   - "@genspark/" npm dependencies — keys in any package.json dependency
 *     block, lockfile entries, and import/require/spawn module specifiers
 *
 * It deliberately does NOT police brand strings: user-facing i18n strings
 * still name upstream trademarks and are swept separately (phase 4b), so
 * i18n directories and documentation files are out of scope here. The
 * license gate lives in tools/check-licenses.mjs; the two run side by side
 * in CI.
 */
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join, dirname, relative, sep } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const SELF = fileURLToPath(import.meta.url)

/** source trees scanned for functional leftovers */
const SCAN_ROOTS = ['apps', 'packages', 'tools', 'scripts']

/** extra root-level manifests checked for @genspark/ dependencies */
const SCAN_FILES = ['package.json', 'package-lock.json']

/** directories never descended into (build output, deps, phase-4b i18n) */
const SKIP_DIRS = new Set([
  'node_modules',
  'dist',
  'out',
  'release',
  'i18n', // brand strings stay until the phase-4b sweep, by decision
])

/** file extensions whose content is scanned (text source and config only) */
const SCAN_EXTS = new Set([
  '.ts',
  '.tsx',
  '.js',
  '.jsx',
  '.cjs',
  '.mjs',
  '.json',
  '.rs',
  '.html',
  '.css',
  '.yml',
  '.yaml',
  '.toml',
])

const MAX_FILE_BYTES = 1_000_000 // skip generated/minified blobs

/** one entry per finding: which rule fired where */
const RULES = [
  {
    name: 'Genspark network domain (genspark.ai / genspark.com)',
    pattern: /genspark\.(ai|com)(?![\w-])/i,
  },
  {
    name: 'retired Genspark login endpoint (device-code / api_tokens / office_addin_auth)',
    pattern: /api_tokens|device[-_]?code|office_addin_auth/i,
  },
  {
    name: '"@genspark/" dependency or module specifier',
    pattern: /@genspark\//,
  },
]

/** recursively collect scannable files under dir */
function collectFiles(dir, out) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (SKIP_DIRS.has(entry.name)) continue
      collectFiles(join(dir, entry.name), out)
    } else if (entry.isFile()) {
      if (!SCAN_EXTS.has(extOf(entry.name))) continue
      const path = join(dir, entry.name)
      if (statSync(path).size > MAX_FILE_BYTES) continue
      if (path === SELF) continue // this file carries the patterns themselves
      out.push(path)
    }
  }
}

function extOf(name) {
  const dot = name.lastIndexOf('.')
  return dot === -1 ? '' : name.slice(dot)
}

const files = []
for (const root of SCAN_ROOTS) {
  collectFiles(join(ROOT, root), files)
}
for (const file of SCAN_FILES) {
  const path = join(ROOT, file)
  if (statSync(path).isFile()) files.push(path)
}

const violations = []
for (const path of files) {
  const text = readFileSync(path, 'utf8')
  const lines = text.split('\n')
  for (const rule of RULES) {
    for (let i = 0; i < lines.length; i++) {
      if (rule.pattern.test(lines[i])) {
        violations.push(`${relative(ROOT, path).split(sep).join('/')}:${i + 1}: ${rule.name}`)
      }
    }
  }
}

if (violations.length > 0) {
  console.error('Functional Genspark leftovers found (egress / dependencies):\n')
  for (const v of violations) console.error(`  ${v}`)
  console.error(
    '\nGenspark network access and dependencies were removed in phase 4a and must not ' +
      'come back. Brand-only strings live in i18n and docs (phase 4b) and are not checked here.\n' +
      'Rules live in tools/check-no-genspark.mjs.',
  )
  process.exit(1)
}

console.log(
  `No Genspark egress leftovers in apps/packages/tools/scripts (${files.length} files scanned).`,
)

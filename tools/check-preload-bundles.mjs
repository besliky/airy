#!/usr/bin/env node
// Guard the sandboxed preload bundles after a build.
//
// Every app preload (apps/<app>/out/preload/index.js) runs inside Electron's
// sandboxed renderer, where `require` resolves only the `electron` module —
// any other external import crashes the whole preload at launch and the
// renderer silently loses every `window.<app>` API (e2e shards then time out
// with no console clue). Regression wave 44 / JOURNAL #203 (PAR-204): a
// `value-import` from shared/desktop-api pulled zod into the sheets preload —
// "module not found: zod", dead preload, 222.88 kB bundle; fixed to 127.10 kB
// by moving the validation into main. The wire-coverage tripwire cannot catch
// this class of bug: it checks which channels exist, not what the bundle
// pulls in or how big it got.
//
// This gate parses each built bundle for:
//   1. external `require` / `import` / dynamic `import()` specifiers — only
//      `electron` is allowed (relative chunks are fine, they are bundled);
//   2. bundle size — 300 kB ceiling with generous headroom over the current
//      ~9-127 kB range, so a dependency sneaking back in trips it.
//
// Run after `npm run build:all` (the bundles must exist): `npm run
// check:preload`. A CI-plan job can call the same script; wiring it into the
// GitHub Actions workflow is a deliberate follow-up. Exit code 1 on any
// violation.
import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

/** Ceiling in kB (SI, 1000 bytes) — current bundles sit at ~9-127 kB. */
export const MAX_PRELOAD_KB = 300

/** The only external module a sandboxed preload may load. */
export const ALLOWED_EXTERNALS = new Set(['electron'])

const APPS_DIR = 'apps'
const BUNDLE_PATH = join('out', 'preload', 'index.js')

// External specifiers in vite/rollup output are emitted as string literals:
//   CJS interop:      const zod = require("zod");
//   ESM output:       import { z } from "zod";  /  import "zod";
//   dynamic import:   const m = await import("zod");
// Lookbehind keeps `foo.require(` / `x.import` out of the matches.
const REQUIRE_RE = /(?<![.\w$])require\(\s*(["'])([^"'\n]+?)\1\s*\)/g
const STATIC_IMPORT_RE = /\bimport\s+(?:[\w*{},$\s]+?\s+from\s+)?(["'])([^"'\n]+?)\1/g
const DYNAMIC_IMPORT_RE = /(?<![.\w$])import\(\s*(["'])([^"'\n]+?)\1\s*\)/g

/**
 * Bare module name of a specifier, or null when it is file-relative
 * (bundled chunks, not an external dependency).
 * 'zod' -> 'zod', 'zod/v4' -> 'zod', '@scope/pkg/sub' -> '@scope/pkg',
 * './chunk-x.js' -> null.
 */
export function bareModuleOf(specifier) {
  if (specifier.startsWith('.') || specifier.startsWith('/')) return null
  const parts = specifier.split('/')
  if (specifier.startsWith('@')) {
    return parts.length >= 2 ? `${parts[0]}/${parts[1]}` : specifier
  }
  return parts[0]
}

/** External (non-electron, non-relative) module names referenced by the source. */
export function parseExternalModules(source) {
  const externals = new Set()
  for (const re of [REQUIRE_RE, STATIC_IMPORT_RE, DYNAMIC_IMPORT_RE]) {
    re.lastIndex = 0
    for (let match = re.exec(source); match; match = re.exec(source)) {
      const bare = bareModuleOf(match[2])
      if (bare !== null && !ALLOWED_EXTERNALS.has(bare)) externals.add(bare)
    }
  }
  return [...externals].sort()
}

/** Human-readable kB size, e.g. "127.1 kB". */
export function formatKb(bytes) {
  return `${(bytes / 1000).toFixed(1)} kB`
}

/**
 * All violations of one bundle: forbidden externals plus size overage.
 * Returns an array of single-line messages (empty when the bundle is clean).
 */
export function checkBundle(source, sizeBytes) {
  const violations = []
  const externals = parseExternalModules(source)
  if (externals.length > 0) {
    violations.push(
      `external module(s) in the sandboxed preload: ${externals.join(', ')} — only ` +
        `'${[...ALLOWED_EXTERNALS].join(', ')}' can be required there. Move the import ` +
        '(e.g. zod or a node built-in) into the main process — see JOURNAL wave 44 #203.',
    )
  }
  if (sizeBytes > MAX_PRELOAD_KB * 1000) {
    violations.push(
      `bundle is ${formatKb(sizeBytes)} (ceiling ${MAX_PRELOAD_KB} kB). Something heavy got ` +
        'pulled into the preload — remove it from preload (see JOURNAL wave 44 #203) or ' +
        'justify a ceiling bump in the PR.',
    )
  }
  return violations
}

/** Every app directory that has a built preload bundle. */
export function findPreloadBundles(root) {
  const appsDir = join(root, APPS_DIR)
  const bundles = []
  for (const entry of readdirSync(appsDir, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue
    const artifact = join(appsDir, entry.name, BUNDLE_PATH)
    if (existsSync(artifact)) {
      bundles.push({ app: entry.name, path: artifact, size: statSync(artifact).size })
    }
  }
  return bundles.sort((a, b) => a.app.localeCompare(b.app))
}

function main() {
  const root = join(fileURLToPath(import.meta.url), '..', '..')
  const bundles = findPreloadBundles(root)
  if (bundles.length === 0) {
    console.error(
      'check-preload-bundles: no apps/<app>/out/preload/index.js found. ' +
        'Run `npm run build:all` first — this gate checks build artifacts.',
    )
    process.exit(1)
  }

  const failures = []
  for (const bundle of bundles) {
    const violations = checkBundle(readFileSync(bundle.path, 'utf8'), bundle.size)
    if (violations.length === 0) {
      console.log(`  ${bundle.app}: OK (${formatKb(bundle.size)})`)
    } else {
      console.error(`  ${bundle.app}: ${bundle.path} (${formatKb(bundle.size)})`)
      for (const message of violations) console.error(`    - ${message}`)
      failures.push(bundle.app)
    }
  }

  if (failures.length > 0) {
    console.error(
      `\ncheck-preload-bundles: FAILED for ${failures.join(', ')}.\n` +
        'Preloads run in Electron\u2019s sandboxed renderer: keep them dependency-free and small. ' +
        'Remove the zod/node import from preload (validate in main instead) — see JOURNAL wave 44 #203.',
    )
    process.exit(1)
  }
  console.log(`check-preload-bundles: OK (${bundles.length} bundles, ceiling ${MAX_PRELOAD_KB} kB)`)
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  main()
}

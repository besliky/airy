// Unit tests for tools/check-preload-bundles.mjs (run: `node --test tools/`).
// They exercise the parser and the verdict logic on synthetic mini-bundles —
// no real build needed: a clean bundle, bundles that require/import an
// external package (the wave 44 #203 zod regression class), and size ceiling
// cases.
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  ALLOWED_EXTERNALS,
  MAX_PRELOAD_KB,
  bareModuleOf,
  checkBundle,
  findPreloadBundles,
  formatKb,
  parseExternalModules,
} from './check-preload-bundles.mjs'

const CLEAN_BUNDLE = [
  '"use strict";',
  'const electron = require("electron");',
  "const chunk = require('./chunk-DEADBEEF.js');",
  'exports.selectWorkbook = () => electron.ipcRenderer.invoke("workbook:select");',
].join('\n')

test('clean bundle: only electron and relative chunks pass', () => {
  assert.deepEqual(parseExternalModules(CLEAN_BUNDLE), [])
  assert.deepEqual(checkBundle(CLEAN_BUNDLE, 1000), [])
})

test('require("zod") is flagged by name (the PAR-204 regression shape)', () => {
  const source = `"use strict";\nconst zod = require("zod");\nconst electron = require("electron");`
  assert.deepEqual(parseExternalModules(source), ['zod'])
  const violations = checkBundle(source, 1000)
  assert.equal(violations.length, 1)
  assert.match(violations[0], /zod/)
  assert.match(violations[0], /JOURNAL wave 44 #203/)
})

test('static ESM import of an external package is flagged', () => {
  const source = `import { z } from "zod";\nexport const s = z.string();`
  assert.deepEqual(parseExternalModules(source), ['zod'])
})

test('bare side-effect import is flagged', () => {
  const source = `import "some-polyfill";\nconst x = 1;`
  assert.deepEqual(parseExternalModules(source), ['some-polyfill'])
})

test('dynamic import() of an external package is flagged', () => {
  const source = `const m = await import("zod");`
  assert.deepEqual(parseExternalModules(source), ['zod'])
})

test('node built-ins are flagged, scoped packages keep scope in the name', () => {
  assert.deepEqual(parseExternalModules('require("node:fs")'), ['node:fs'])
  assert.deepEqual(parseExternalModules('require("@scope/secret/pkg")'), ['@scope/secret'])
})

test('member access and identifiers named require/import are not matches', () => {
  const source = `foo.require("zod");\nx.import("zod");\nmyrequire("zod");\nconst importer = 1;`
  assert.deepEqual(parseExternalModules(source), [])
})

test('size ceiling: 301 kB fails, 299 kB passes, message names the ceiling', () => {
  const big = CLEAN_BUNDLE + '/*' + 'x'.repeat(301_000) + '*/'
  const sizeViolations = checkBundle(CLEAN_BUNDLE, big.length)
  assert.equal(sizeViolations.length, 1)
  assert.match(sizeViolations[0], /ceiling/)
  assert.deepEqual(checkBundle(CLEAN_BUNDLE, (MAX_PRELOAD_KB - 1) * 1000), [])
})

test('external and size violations are reported together', () => {
  const violations = checkBundle(`require("zod");`, (MAX_PRELOAD_KB + 1) * 1000)
  assert.equal(violations.length, 2)
})

test('bareModuleOf reduces specifiers and drops file-relative ones', () => {
  assert.equal(bareModuleOf('zod'), 'zod')
  assert.equal(bareModuleOf('zod/v4'), 'zod')
  assert.equal(bareModuleOf('@scope/pkg'), '@scope/pkg')
  assert.equal(bareModuleOf('@scope/pkg/sub'), '@scope/pkg')
  assert.equal(bareModuleOf('./chunk-x.js'), null)
  assert.equal(bareModuleOf('/abs/path.js'), null)
})

test('formatKb renders SI kilobytes with one decimal', () => {
  assert.equal(formatKb(127_100), '127.1 kB')
  assert.equal(formatKb(9331), '9.3 kB')
})

test('electron is the only allowed external', () => {
  assert.deepEqual([...ALLOWED_EXTERNALS].sort(), ['electron'])
})

test('findPreloadBundles discovers only apps with a built artifact, sorted', async () => {
  const root = await mkdtemp(join(tmpdir(), 'preload-guard-'))
  try {
    for (const app of ['sheets', 'docs']) {
      const dir = join(root, 'apps', app, 'out', 'preload')
      await mkdir(dir, { recursive: true })
      await writeFile(join(dir, 'index.js'), CLEAN_BUNDLE)
    }
    // An app without a build output must be skipped, not reported.
    await mkdir(join(root, 'apps', 'empty'), { recursive: true })
    // Non-app directories are ignored.
    await mkdir(join(root, 'packages', 'ui'), { recursive: true })

    const bundles = findPreloadBundles(root)
    assert.deepEqual(
      bundles.map((b) => b.app),
      ['docs', 'sheets'],
    )
    assert.equal(bundles[0].size, Buffer.byteLength(CLEAN_BUNDLE))
    assert.equal(bundles[0].path, join(root, 'apps', 'docs', 'out', 'preload', 'index.js'))
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

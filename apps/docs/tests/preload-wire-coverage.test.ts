import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { BRIDGE_INVOKE_CHANNEL, BRIDGE_RESULT_CHANNEL } from '../src/shared/ipc'

/**
 * Tripwire for the preload-wire contract of the docs app (TEST-604, modeled
 * after apps/sheets/tests/preload-wire-coverage.test.ts).
 *
 * Unlike sheets, the docs preload is a thin pass-through: it validates no
 * payload, so there is no runtime whitelist to cross-check against a zod
 * schema. The equivalent trap here is contract drift that typecheck cannot
 * see: a channel string typo in a hand-written literal, or a member silently
 * vanishing from the shared DesktopApi interface together with its preload
 * wiring (compiles clean, renderer just loses the capability).
 *
 * These tests pin the wire surface so that any shrink, rename, or unwired
 * addition of the shared schema is a deliberate, reviewable change:
 * - the shared DesktopApi interface must still declare the pinned members;
 * - the preload api object must implement exactly that interface surface;
 * - every channel literal the preload consumes must be produced by a
 *   main-process source (docs-main or the shell), so typos redden here
 *   instead of failing silently at runtime.
 */
const here = dirname(fileURLToPath(import.meta.url))
const preloadSource = readFileSync(join(here, '../src/preload/index.ts'), 'utf8')
const sharedSource = readFileSync(join(here, '../src/shared/ipc.ts'), 'utf8')

/** Member names of `export interface <name>` in the shared schema source. */
function interfaceMembers(source: string, name: string): string[] {
  const start = source.search(new RegExp(`export interface ${name}\\b`))
  if (start < 0) throw new Error(`shared schema no longer declares ${name}`)
  const body = source.slice(start, source.indexOf('\n}', start))
  const members: string[] = []
  for (const line of body.split('\n')) {
    const match = line.match(/^ {2}([A-Za-z_$][\w$]*)\s*\??\s*[:<(]/)
    if (match) members.push(match[1])
  }
  return members
}

/** Top-level keys of the hand-written preload api object. */
function apiObjectKeys(source: string, declaration: string): string[] {
  const start = source.indexOf(declaration)
  if (start < 0) throw new Error(`preload no longer declares "${declaration}"`)
  const body = source.slice(start, source.indexOf('\n}', start))
  const keys: string[] = []
  for (const line of body.split('\n')) {
    const match = line.match(/^ {2}([A-Za-z_$][\w$]*):/)
    if (match && !['const', 'let', 'return'].includes(match[1])) keys.push(match[1])
  }
  return keys
}

function collectSources(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name)
    if (entry.isDirectory()) collectSources(path, out)
    else if (entry.name.endsWith('.ts')) out.push(readFileSync(path, 'utf8'))
  }
  return out
}

// The docs tab's own main-process registration plus the shell's app-wide
// handlers (app:*, menu:*, and the bridge server under src/main/bridge).
const mainCorpus = [
  ...collectSources(join(here, '../src/main')),
  ...collectSources(join(here, '../../shell/src/main')),
].join('\n')

/** Frozen surface of window.desktopApi — update consciously when it grows. */
const PINNED_API_MEMBERS = `
  addAttachmentPaths
  addPastedImage
  aiChat
  aiGenerateImage
  aiStream
  aiStreamCancel
  consumeAiDocContent
  consumeNewBlankDoc
  consumePendingOpenDocx
  copyImageToClipboard
  createDocument
  discardDocPasswordIntents
  docPasswordIntentRevision
  exportHtml
  exportPdf
  fetchImage
  focusDocsTab
  fontMetrics
  getAiPanelPrefs
  getAiSettings
  getAuthorName
  getAutoSaveDefault
  getLanguage
  getPathForFile
  getRecentFiles
  getTheme
  imageSearch
  listDocsTabs
  onAiPanelPrefsChanged
  onAiStream
  onAuthorNameChanged
  onAutoSaveDefaultChanged
  onBridgeInvoke
  onChromePressed
  onCloseCheck
  onCloseSaveRequest
  onLanguageChanged
  onMenuCommand
  onOpenDocx
  onRenamedDocx
  onTeardown
  onThemeChanged
  openDocx
  openDocxDecrypt
  openDocxPath
  openNewTab
  pickAttachments
  pickImage
  print
  printPdfBuffer
  readAttachment
  readAttachmentImage
  reportBridgeResult
  reportCloseCheck
  reportCloseSaveResult
  reportViewMenuState
  respellKick
  saveDocx
  saveDocxAs
  saveDocxNew
  saveMergedPdf
  setAiSettings
  setDocPassword
  webSearch
  writeRecoveryCopy
`
  .trim()
  .split(/\s+/)
  .sort()

/** Frozen fields of the open-file wire payload (main -> renderer). */
const PINNED_OPEN_FILE_FIELDS = `
  data
  encrypted
  hash
  name
  path
  recovered
`
  .trim()
  .split(/\s+/)
  .sort()

describe('docs preload wire coverage', () => {
  it('shared DesktopApi still declares the pinned renderer surface', () => {
    const members = interfaceMembers(sharedSource, 'DesktopApi')
    expect(members.length).toBeGreaterThan(60)
    expect([...members].sort()).toEqual(PINNED_API_MEMBERS)
  })

  it('preload implements exactly the shared DesktopApi surface', () => {
    const members = interfaceMembers(sharedSource, 'DesktopApi')
    const wired = apiObjectKeys(preloadSource, 'const api: DesktopApi = {')
    expect(wired.length).toBe(members.length)
    expect([...wired].sort()).toEqual([...members].sort())
  })

  it('every channel literal the preload consumes is produced by main', () => {
    const literals = [...new Set(preloadSource.match(/'([a-z][\w-]*:[\w-]+)'/g) ?? [])]
    // Guard against the extraction rotting into a vacuous pass.
    expect(literals.length).toBeGreaterThan(60)
    const orphaned = literals.filter((quoted) => !mainCorpus.includes(quoted))
    expect(orphaned).toEqual([])
  })

  it('shell bridge keeps wiring the shared bridge channels', () => {
    // The preload and the shell bridge must keep importing the channel names
    // from the same shared module; a hand-copied literal here is how the
    // bridge silently stops reaching the renderer.
    expect(preloadSource).toContain('BRIDGE_INVOKE_CHANNEL')
    expect(preloadSource).toContain('BRIDGE_RESULT_CHANNEL')
    expect(mainCorpus).toContain('BRIDGE_INVOKE_CHANNEL')
    expect(mainCorpus).toContain('BRIDGE_RESULT_CHANNEL')
    expect([BRIDGE_INVOKE_CHANNEL, BRIDGE_RESULT_CHANNEL]).toEqual([
      'airy-bridge:invoke',
      'airy-bridge:result',
    ])
  })

  it('pins the OpenFileResult wire fields', () => {
    const fields = interfaceMembers(sharedSource, 'OpenFileResult')
    expect([...fields].sort()).toEqual(PINNED_OPEN_FILE_FIELDS)
  })
})

import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { AI_CHANNELS, HTML_CHANNELS } from '../src/shared/ipc'

/**
 * Tripwire for the preload-wire contract of the html app (TEST-604, modeled
 * after apps/sheets/tests/preload-wire-coverage.test.ts).
 *
 * The html preload is a thin pass-through: it validates no payload, so there
 * is no whitelist to cross-check against a zod schema. The html flavor of the
 * sheets trap is channel drift against the shared channel registry: HTML_/AI_
 * CHANNELS gain an entry (or SaveHtmlRequest loses a field) while the
 * hand-written preload keeps wiring the old surface — everything compiles,
 * the renderer just never reaches the new wire.
 *
 * These tests pin the wire surface so that any shrink, rename, or unwired
 * addition of the shared schema is a deliberate, reviewable change:
 * - the shared HtmlApi interface and the channel registries must still
 *   declare the pinned entries;
 * - the preload api object must implement exactly that interface surface;
 * - every registry channel must be referenced by the preload (the direct
 *   analog of "the whitelist must name every schema field");
 * - the few raw channel literals (project:*, app:chrome-pressed) must be
 *   produced by a main-process source, so typos redden here.
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

// html:* handlers live in html main; app:chrome-pressed and the shared ai:*
// handlers are registered by the shell; the project:* chat-persistence
// handlers are registered centrally by docs-main's registerProjectIpc.
const mainCorpus = [
  ...collectSources(join(here, '../src/main')),
  ...collectSources(join(here, '../../shell/src/main')),
  readFileSync(join(here, '../../docs/src/main/docs-main.ts'), 'utf8'),
].join('\n')

/** Frozen surface of window.htmlApi — update consciously when it grows. */
const PINNED_API_MEMBERS = `
  addAttachmentPaths
  addPastedImage
  aiGenerateImage
  aiStream
  aiStreamCancel
  consumePending
  exportDocx
  exportPdf
  fetchImage
  getAiPanelPrefs
  getAiSettings
  getAutoSaveDefault
  getEditorPrefs
  getLanguage
  getPathForFile
  getPreviewInfo
  getTheme
  imageSearch
  onAiPanelPrefsChanged
  onAiStream
  onAutoSaveDefaultChanged
  onChromePressed
  onCloseSaveRequest
  onExportRequest
  onFileRenamed
  onLanguageChanged
  onPrintRequest
  onSaveRequest
  onThemeChanged
  pickAttachments
  pickImage
  presentInNewTab
  readAttachment
  readAttachmentImage
  readFile
  readImage
  save
  saveImage
  sendCloseSaveResult
  sendSaveRequestAck
  setDirty
  setEditorPrefs
  setEncoding
  setPresentFullScreen
  setProvisionalTitle
  updatePreview
  webSearch
  writeRecovery
`
  .trim()
  .split(/\s+/)
  .sort()

/** Frozen html channel registry — must match HTML_CHANNELS key for key. */
const PINNED_HTML_CHANNELS = `
  aiGenerateImage
  aiPanelPrefsChanged
  autoSaveDefaultChanged
  closeSaveRequest
  closeSaveResult
  consumePending
  dirtyChanged
  exportDocx
  exportPdf
  exportRequest
  fetchImage
  fileRenamed
  filesAdd
  filesAddPastedImage
  filesPick
  filesRead
  filesReadImage
  getEditorPrefs
  getAiPanelPrefs
  getAutoSaveDefault
  getLanguage
  getTheme
  languageChanged
  pickImage
  presentFullScreen
  presentNewTab
  previewInfo
  previewUpdate
  printRequest
  provisionalTitle
  readFile
  readImage
  save
  saveImage
  saveRequest
  saveRequestAck
  setEditorPrefs
  setEncoding
  themeChanged
  writeRecovery
`
  .trim()
  .split(/\s+/)
  .sort()

/** Frozen shared-AI channel registry used by the html preload. */
const PINNED_AI_CHANNELS = `
  getSettings
  imageSearch
  stream
  streamCancel
  streamChunk
  webSearch
`
  .trim()
  .split(/\s+/)
  .sort()

/** Frozen fields of the save-request wire payload (renderer -> main). */
const PINNED_SAVE_REQUEST_FIELDS = `
  auto
  defaultName
  imageSources
  mode
  suggestedName
  text
`
  .trim()
  .split(/\s+/)
  .sort()

describe('html preload wire coverage', () => {
  it('shared HtmlApi still declares the pinned renderer surface', () => {
    const members = interfaceMembers(sharedSource, 'HtmlApi')
    expect(members.length).toBeGreaterThan(40)
    expect([...members].sort()).toEqual(PINNED_API_MEMBERS)
  })

  it('preload implements exactly the shared HtmlApi surface', () => {
    const members = interfaceMembers(sharedSource, 'HtmlApi')
    const wired = apiObjectKeys(preloadSource, 'const api: HtmlApi = {')
    expect(wired.length).toBe(members.length)
    expect([...wired].sort()).toEqual([...members].sort())
  })

  it('channel registries still declare the pinned entries', () => {
    expect(Object.keys(HTML_CHANNELS).sort()).toEqual(PINNED_HTML_CHANNELS)
    expect(Object.keys(AI_CHANNELS).sort()).toEqual(PINNED_AI_CHANNELS)
  })

  it('preload wires every entry of both channel registries', () => {
    // A registry entry the preload never references is the html flavor of the
    // sheets trap: the shared contract grew, the hand-written wire did not.
    const unwired = [
      ...Object.keys(HTML_CHANNELS)
        .filter((key) => !new RegExp(`HTML_CHANNELS\\.${key}\\b`).test(preloadSource))
        .map((key) => `HTML_CHANNELS.${key}`),
      ...Object.keys(AI_CHANNELS)
        .filter((key) => !new RegExp(`AI_CHANNELS\\.${key}\\b`).test(preloadSource))
        .map((key) => `AI_CHANNELS.${key}`),
    ]
    expect(unwired).toEqual([])
  })

  it('every raw channel literal the preload consumes is produced by main', () => {
    const literals = [...new Set(preloadSource.match(/'([a-z][\w-]*:[\w-]+)'/g) ?? [])]
    expect(literals.length).toBeGreaterThan(4)
    const orphaned = literals.filter((quoted) => !mainCorpus.includes(quoted))
    expect(orphaned).toEqual([])
  })

  it('pins the SaveHtmlRequest wire fields', () => {
    const fields = interfaceMembers(sharedSource, 'SaveHtmlRequest')
    expect([...fields].sort()).toEqual(PINNED_SAVE_REQUEST_FIELDS)
  })
})

import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

/**
 * Tripwire for the preload-wire contract of the slides app (TEST-604, modeled
 * after apps/sheets/tests/preload-wire-coverage.test.ts).
 *
 * The slides preload is a thin pass-through with no payload validation and
 * — unlike pdf/html/markdown — no shared channel constants: every channel is
 * a hand-written string literal on both sides of the wire. The equivalent of
 * the sheets whitelist trap is therefore literal drift: a typo'd or renamed
 * channel compiles clean and silently breaks the renderer<->main wire.
 *
 * These tests pin the wire surface so that any shrink, rename, or unwired
 * change is a deliberate, reviewable edit:
 * - the shared SlidesApi/DesktopFilesApi interfaces must still declare the
 *   pinned members;
 * - the preload objects must implement exactly those interface surfaces;
 * - every channel literal the preload consumes must be produced by a
 *   main-process source (slides main, the shell, or docs-main, which
 *   centrally registers the shared project:* handlers in shell mode).
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

/** Top-level keys of a hand-written preload api object. */
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

// slides:* and its own app:* handlers live in slides main; app:chrome-pressed
// and the shared ai:* handlers are registered by the shell; the project:*
// chat-persistence handlers are registered centrally by docs-main's
// registerProjectIpc in shell mode (slides standalone covers the 4 chat
// channels only), so that source is part of the wire too.
const mainCorpus = [
  ...collectSources(join(here, '../src/main')),
  ...collectSources(join(here, '../../shell/src/main')),
  readFileSync(join(here, '../../docs/src/main/docs-main.ts'), 'utf8'),
].join('\n')

/** Frozen surface of window.slidesApi — update consciously when it grows. */
const PINNED_API_MEMBERS = `
  addBlankSlide
  addChart
  addComment
  addElement
  addImageBytes
  addInk
  addMediaBytes
  addSection
  addSlide
  addSlideWithLayout
  addSmartArt
  addTable
  aiLogRunFailure
  aiSnapshotRestore
  aiStream
  aiStreamCancel
  analyzeMedia
  applyEditScript
  applyHeaderFooter
  applyTheme
  applyTxn
  audienceNav
  audienceReady
  batchEditTransform
  beginHistoryBatch
  changeShape
  clipboardExternal
  clipboardProbe
  consumePendingOpen
  copyElements
  copySlide
  deleteComment
  deleteElement
  deleteSlide
  duplicateElements
  editBackground
  editChart
  editConnectorEndpoints
  editFill
  editImageFill
  editPictureOpacity
  editPictureSrcRect
  editStroke
  editTableCell
  editTableStyle
  editText
  editTransform
  endHistoryBatch
  exportImages
  exportPdf
  findReplace
  flipElements
  fontCatalog
  fontDownload
  fontInstallLocal
  fontMissing
  generateImage
  getAiPanelPrefs
  getAiSettings
  getAnimations
  getAutoSaveDefault
  getChartColorSchemes
  getChartData
  getComments
  getHeaderFooter
  getLanguage
  getLayouts
  getLink
  getMediaData
  getNotes
  getRecentFiles
  getRenderSlides
  getRunLinks
  getSections
  getShapeKeys
  getSlideLinks
  getSlideSize
  getTheme
  getTransition
  groupElements
  hasSlideClipboard
  imageSearch
  insertImage
  insertImageUrl
  insertMedia
  insertModel3d
  isDirty
  landGeneratedPages
  listStyleTemplates
  loadStyleTemplate
  localGeneratePage
  masterClose
  masterDeleteElement
  masterEditFill
  masterEditStroke
  masterEditText
  masterEditTransform
  masterEnter
  masterOpen
  moveSection
  moveSlide
  nativeClipboard
  newBlank
  onAiPanelPrefsChanged
  onAiStream
  onAudienceNav
  onAutoSaveDefaultChanged
  onChromePressed
  onCloseSaveRequest
  onDeckChanged
  onFontsChanged
  onHistoryChanged
  onLanguageChanged
  onMenuCommand
  onOpened
  onRenamed
  onShowInk
  onShowSync
  onThemeChanged
  openPptx
  openPptxPath
  pasteElements
  pasteSlide
  pickExportDir
  pickExportPdfPath
  pickPictureFile
  presenterEnd
  presenterInk
  presenterStart
  presenterSwap
  presenterSync
  printSlides
  privateFontData
  privateFontFaces
  redo
  removeSection
  renameSection
  reorderElement
  repasteSlide
  replacePictureBytes
  replacePictureUrl
  reportCloseSaveResult
  resolveComments
  save
  saveAs
  saveStyleSidecar
  saveStyleTemplate
  setAdvanceTimes
  setAiSettings
  setAnimations
  setAutoSavePref
  setEffects
  setElementFont
  setElementParagraphFormat
  setLink
  setNotes
  setSections
  setShapeAdjust
  setShowFullScreen
  setSlideHidden
  setSlideLayout
  setSlideSize
  setTableCellAnchor
  setTableColWidth
  setTableRowHeight
  setTextAnchor
  setTextBodyProps
  setTransition
  tableMerge
  tableStructure
  undo
  ungroupElement
  webSearch
`
  .trim()
  .split(/\s+/)
  .sort()

/** Frozen surface of the window.desktop files-attachment subset. */
const PINNED_FILES_API_MEMBERS = `
  addAttachmentPaths
  addPastedImage
  getPathForFile
  pickAttachments
  readAttachment
  readAttachmentImage
`
  .trim()
  .split(/\s+/)
  .sort()

/** Frozen fields of the open-deck wire payload (main -> renderer). */
const PINNED_OPEN_RESULT_FIELDS = `
  defaultFont
  path
  size
  slides
`
  .trim()
  .split(/\s+/)
  .sort()

describe('slides preload wire coverage', () => {
  it('shared SlidesApi still declares the pinned renderer surface', () => {
    const members = interfaceMembers(sharedSource, 'SlidesApi')
    expect(members.length).toBeGreaterThan(160)
    expect([...members].sort()).toEqual(PINNED_API_MEMBERS)
  })

  it('preload implements exactly the shared SlidesApi surface', () => {
    const members = interfaceMembers(sharedSource, 'SlidesApi')
    const wired = apiObjectKeys(preloadSource, 'const api: SlidesApi = {')
    expect(wired.length).toBe(members.length)
    expect([...wired].sort()).toEqual([...members].sort())
  })

  it('files bridge implements exactly the shared DesktopFilesApi surface', () => {
    const members = interfaceMembers(sharedSource, 'DesktopFilesApi')
    expect([...members].sort()).toEqual(PINNED_FILES_API_MEMBERS)
    const wired = apiObjectKeys(preloadSource, 'const filesApi: DesktopFilesApi = {')
    expect([...wired].sort()).toEqual(PINNED_FILES_API_MEMBERS)
  })

  it('every channel literal the preload consumes is produced by main', () => {
    const literals = [...new Set(preloadSource.match(/'([a-z][\w-]*:[\w-]+)'/g) ?? [])]
    // Guard against the extraction rotting into a vacuous pass.
    expect(literals.length).toBeGreaterThan(170)
    const orphaned = literals.filter((quoted) => !mainCorpus.includes(quoted))
    expect(orphaned).toEqual([])
  })

  it('pins the OpenResult wire fields', () => {
    const fields = interfaceMembers(sharedSource, 'OpenResult')
    expect([...fields].sort()).toEqual(PINNED_OPEN_RESULT_FIELDS)
  })
})

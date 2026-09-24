import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * BUG-1675: a Home rename moved the file on disk but only re-pointed tabs in
 * the FOCUSED window — a second window's tab kept the stale title/save path,
 * so Ctrl+S there wrote to the pre-rename name (caught only by the on-disk
 * staleness fence). The broadcast now runs through renameFileInAllWindows,
 * which walks EVERY window's TabManager.
 *
 * Behavioral half: two real TabManagers (two windows); the rename must update
 * matching tabs in both and notify each affected editor exactly once.
 */

interface FakeWebContents {
  id: number
  on: ReturnType<typeof vi.fn>
  removeListener: ReturnType<typeof vi.fn>
  close: ReturnType<typeof vi.fn>
  reload: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
}

interface FakeView {
  webContents: FakeWebContents
  setVisible: ReturnType<typeof vi.fn>
  setBounds: ReturnType<typeof vi.fn>
}

let nextWebContentsId = 1

function makeFakeView(): FakeView {
  return {
    webContents: {
      id: nextWebContentsId++,
      on: vi.fn(),
      removeListener: vi.fn(),
      close: vi.fn(),
      reload: vi.fn(),
      loadURL: vi.fn(() => Promise.resolve()),
      isDestroyed: vi.fn(() => false),
    },
    setVisible: vi.fn(),
    setBounds: vi.fn(),
  }
}

vi.mock('electron', () => ({
  BrowserWindow: class {},
  // markdown/html module-level recovery stores resolve their userData dir
  app: { getPath: () => '/tmp/airy-rename-broadcast-test' },
}))

const createDocsView = vi.fn(() => makeFakeView())

vi.mock('../../docs/src/main/docs-main', () => ({
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: vi.fn(() => Promise.resolve(false)),
  markDocsNewBlank: vi.fn(),
  recordRecentFile: vi.fn(),
  requestDocsClose: vi.fn(() => Promise.resolve(true)),
  setActiveDocsResolver: vi.fn(),
  teardownDocsRenderer: vi.fn(),
}))

const createMarkdownView = vi.fn(() => makeFakeView())

vi.mock('../../markdown/src/main/markdown-main', () => ({
  createMarkdownView: (...args: unknown[]) => createMarkdownView(...(args as [])),
  markdownIsDirty: vi.fn(() => false),
  requestMarkdownClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: vi.fn(() => makeFakeView()),
  pdfIsDirty: vi.fn(() => false),
  clearPdfDirty: vi.fn(),
  requestPdfClose: vi.fn(() => Promise.resolve(true)),
}))

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: vi.fn(() => makeFakeView()),
  queueWorkbookForView: vi.fn(),
  requestSheetsClose: vi.fn(() => Promise.resolve(true)),
  setActiveSheetsWebContents: vi.fn(),
  setSheetsNewBlank: vi.fn(),
  sheetsPendingEditCount: vi.fn(() => 0),
}))

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: vi.fn(() => makeFakeView()),
  requestSlidesClose: vi.fn(() => Promise.resolve(true)),
  setActiveSlidesWebContents: vi.fn(),
  slidesIsDirty: vi.fn(() => false),
}))

import { TabManager } from '../src/main/tab-manager'
import { renameFileInAllWindows, type FileRenamedHooks } from '../src/main/rename-broadcast'

interface FakeShellWindow {
  on: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
  getContentBounds: () => { x: number; y: number; width: number; height: number }
  contentView: {
    addChildView: ReturnType<typeof vi.fn>
    removeChildView: ReturnType<typeof vi.fn>
  }
}

function makeShellWindow(): FakeShellWindow {
  return {
    on: vi.fn(),
    isDestroyed: vi.fn(() => false),
    getContentBounds: () => ({ x: 0, y: 0, width: 800, height: 600 }),
    contentView: { addChildView: vi.fn(), removeChildView: vi.fn() },
  }
}

function makeManager(): TabManager {
  return new TabManager(makeShellWindow() as never, vi.fn(), vi.fn())
}

function makeHooks(): FileRenamedHooks & Record<string, ReturnType<typeof vi.fn>> {
  const fn = () => vi.fn()
  return { docs: fn(), sheets: fn(), slides: fn(), markdown: fn(), html: fn() }
}

describe('renameFileInAllWindows (BUG-1675)', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    nextWebContentsId = 1
  })

  it('re-points matching tabs in BOTH windows, so the second window saves to the new path', () => {
    const windowA = makeManager()
    const windowB = makeManager()
    // the audited scenario: the file lives in window B while window A has focus
    const docsId = windowA.openDocsTab('/tmp/reports/report.docx')
    const staleId = windowB.openMarkdownTab('/tmp/reports/report.docx')

    const hooks = makeHooks()
    renameFileInAllWindows(
      [windowA, windowB],
      '/tmp/reports/report.docx',
      '/tmp/reports/report-renamed.docx',
      hooks,
    )

    for (const manager of [windowA, windowB]) {
      expect(manager.list().find((t) => t.id !== 'home')?.title).toBe('report-renamed.docx')
      expect(manager.findTabIdByPath('docs', '/tmp/reports/report.docx')).toBeUndefined()
      expect(manager.findTabIdByPath('markdown', '/tmp/reports/report.docx')).toBeUndefined()
    }
    // the second window's save path now follows the renamed file — no stale save
    expect(windowB.findTabIdByPath('markdown', '/tmp/reports/report-renamed.docx')).toBe(staleId)
    expect(windowA.findTabIdByPath('docs', '/tmp/reports/report-renamed.docx')).toBe(docsId)
    expect(windowA.list().find((t) => t.id === docsId)?.title).toBe('report-renamed.docx')

    // each affected editor is notified once, with its own webContents
    expect(hooks.docs).toHaveBeenCalledTimes(1)
    expect(hooks.docs).toHaveBeenCalledWith(
      (createDocsView.mock.results.at(-1)!.value as FakeView).webContents,
      '/tmp/reports/report.docx',
      '/tmp/reports/report-renamed.docx',
    )
    expect(hooks.markdown).toHaveBeenCalledTimes(1)
    expect(hooks.markdown).toHaveBeenCalledWith(
      (createMarkdownView.mock.results.at(-1)!.value as FakeView).webContents,
      '/tmp/reports/report.docx',
      '/tmp/reports/report-renamed.docx',
    )
    expect(hooks.sheets).not.toHaveBeenCalled()
    expect(hooks.slides).not.toHaveBeenCalled()
    expect(hooks.html).not.toHaveBeenCalled()
  })

  it('leaves non-matching tabs and windows untouched', () => {
    const windowA = makeManager()
    const windowB = makeManager()
    windowA.openDocsTab('/tmp/other.docx')
    windowB.openMarkdownTab('/tmp/winB.md')

    const hooks = makeHooks()
    renameFileInAllWindows([windowA, windowB], '/tmp/absent.docx', '/tmp/renamed.docx', hooks)

    const otherId = windowA.findTabIdByPath('docs', '/tmp/other.docx')
    expect(otherId).toBeDefined()
    expect(windowA.list().find((t) => t.id === otherId)?.title).toBe('other.docx')
    const winBId = windowB.findTabIdByPath('markdown', '/tmp/winB.md')
    expect(winBId).toBeDefined()
    expect(windowB.list().find((t) => t.id === winBId)?.title).toBe('winB.md')
    for (const hook of Object.values(hooks)) expect(hook).not.toHaveBeenCalled()
  })
})

/**
 * Wiring half: index.ts is an Electron main module that cannot be imported
 * into a unit test, so the routing itself is the contract under regression
 * guard (same style as session-restore-cascade.test.ts).
 */
describe('home rename broadcast wiring', () => {
  const here = dirname(fileURLToPath(import.meta.url))
  const source = readFileSync(join(here, '../src/main/index.ts'), 'utf8')

  it('the home rename handler broadcasts to every window, not the focused one', () => {
    expect(source).toContain('shellWindows.list().map((entry) => entry.manager)')
    expect(source).not.toContain('focusedManager()?.renameTabFile')
  })

  it('routes through the all-windows helper with every per-editor renamed hook', () => {
    expect(source).toContain('renameFileInAllWindows(')
    for (const hook of ['docs', 'sheets', 'slides', 'markdown', 'html']) {
      expect(source).toContain(`${hook}: ${hook}FileRenamed`)
    }
  })
})

import { beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * TabManager (src/main/tab-manager.ts): tab list state, activation,
 * close guards, and view lifecycle inside the shell's single window.
 * Electron and the per-module main entrypoints are mocked; only the
 * manager's own observable behavior is asserted.
 */

interface FakeWebContents {
  id: number
  on: ReturnType<typeof vi.fn>
  removeListener: ReturnType<typeof vi.fn>
  close: ReturnType<typeof vi.fn>
  reload: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
  listeners: Map<string, (event?: unknown, details?: unknown) => void>
}

interface FakeView {
  webContents: FakeWebContents
  setVisible: ReturnType<typeof vi.fn>
  setBounds: ReturnType<typeof vi.fn>
}

let nextWebContentsId = 1

function makeFakeView(): FakeView {
  const listeners = new Map<string, () => void>()
  return {
    webContents: {
      id: nextWebContentsId++,
      listeners,
      on: vi.fn((event: string, handler: () => void) => {
        listeners.set(event, handler)
      }),
      removeListener: vi.fn((event: string, handler: () => void) => {
        if (listeners.get(event) === handler) listeners.delete(event)
      }),
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
  app: { getPath: () => '/tmp/airy-tab-manager-test' },
}))

const createDocsView = vi.fn(() => makeFakeView())
const docsQueryDirty = vi.fn(() => Promise.resolve(false))
const markDocsNewBlank = vi.fn()
const recordRecentFile = vi.fn()
const requestDocsClose = vi.fn(() => Promise.resolve(true))
const setActiveDocsResolver = vi.fn()
const teardownDocsRenderer = vi.fn()

vi.mock('../../docs/src/main/docs-main', () => ({
  createDocsView: (...args: unknown[]) => createDocsView(...(args as [])),
  docsQueryDirty: (...args: unknown[]) => docsQueryDirty(...(args as [])),
  markDocsNewBlank: (...args: unknown[]) => markDocsNewBlank(...args),
  recordRecentFile: (...args: unknown[]) => recordRecentFile(...args),
  requestDocsClose: (...args: unknown[]) => requestDocsClose(...(args as [])),
  setActiveDocsResolver: (...args: unknown[]) => setActiveDocsResolver(...args),
  teardownDocsRenderer: (...args: unknown[]) => teardownDocsRenderer(...args),
}))

const createPdfView = vi.fn(() => makeFakeView())
const pdfIsDirty = vi.fn(() => false)
const clearPdfDirty = vi.fn()
const requestPdfClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../pdf/src/main/pdf-main', () => ({
  createPdfView: (...args: unknown[]) => createPdfView(...(args as [])),
  pdfIsDirty: (...args: unknown[]) => pdfIsDirty(...(args as [])),
  clearPdfDirty: (...args: unknown[]) => clearPdfDirty(...(args as [])),
  requestPdfClose: (...args: unknown[]) => requestPdfClose(...(args as [])),
}))

const createSheetsView = vi.fn(() => makeFakeView())
const queueWorkbookForView = vi.fn()
const requestSheetsClose = vi.fn(() => Promise.resolve(true))
const setActiveSheetsWebContents = vi.fn()
const setSheetsNewBlank = vi.fn()
const sheetsPendingEditCount = vi.fn(() => 0)

vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsView: (...args: unknown[]) => createSheetsView(...(args as [])),
  queueWorkbookForView: (...args: unknown[]) => queueWorkbookForView(...args),
  requestSheetsClose: (...args: unknown[]) => requestSheetsClose(...(args as [])),
  setActiveSheetsWebContents: (...args: unknown[]) => setActiveSheetsWebContents(...args),
  setSheetsNewBlank: (...args: unknown[]) => setSheetsNewBlank(...args),
  sheetsPendingEditCount: (...args: unknown[]) => sheetsPendingEditCount(...(args as [])),
}))

const createSlidesView = vi.fn(() => makeFakeView())
const requestSlidesClose = vi.fn(() => Promise.resolve(true))
const setActiveSlidesWebContents = vi.fn()
const slidesIsDirty = vi.fn(() => false)

vi.mock('../../slides/src/main/slides-main', () => ({
  createSlidesView: (...args: unknown[]) => createSlidesView(...(args as [])),
  requestSlidesClose: (...args: unknown[]) => requestSlidesClose(...(args as [])),
  setActiveSlidesWebContents: (...args: unknown[]) => setActiveSlidesWebContents(...args),
  slidesIsDirty: (...args: unknown[]) => slidesIsDirty(...(args as [])),
}))

const createMarkdownView = vi.fn(() => makeFakeView())
const markdownIsDirty = vi.fn(() => false)
const requestMarkdownClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../markdown/src/main/markdown-main', () => ({
  createMarkdownView: (...args: unknown[]) => createMarkdownView(...(args as [])),
  markdownIsDirty: (...args: unknown[]) => markdownIsDirty(...(args as [])),
  requestMarkdownClose: (...args: unknown[]) => requestMarkdownClose(...(args as [])),
}))

const createHtmlView = vi.fn(() => makeFakeView())
const createHtmlPresentView = vi.fn(() => makeFakeView())
const htmlIsDirty = vi.fn(() => false)
const requestHtmlClose = vi.fn(() => Promise.resolve(true))

vi.mock('../../html/src/main/html-main', () => ({
  createHtmlView: (...args: unknown[]) => createHtmlView(...(args as [])),
  createHtmlPresentView: (...args: unknown[]) => createHtmlPresentView(...(args as [])),
  htmlIsDirty: (...args: unknown[]) => htmlIsDirty(...(args as [])),
  requestHtmlClose: (...args: unknown[]) => requestHtmlClose(...(args as [])),
}))

import { TabManager } from '../src/main/tab-manager'

const TAB_STRIP_HEIGHT = 40
const WINDOW_WIDTH = 800
const WINDOW_HEIGHT = 600

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
    getContentBounds: () => ({ x: 0, y: 0, width: WINDOW_WIDTH, height: WINDOW_HEIGHT }),
    contentView: { addChildView: vi.fn(), removeChildView: vi.fn() },
  }
}

let shellWindow: FakeShellWindow
let onChanged: ReturnType<typeof vi.fn>
let applyMenuFor: ReturnType<typeof vi.fn>
let manager: TabManager

function lastCreatedView(factory: ReturnType<typeof vi.fn>): FakeView {
  return factory.mock.results.at(-1)!.value as FakeView
}

beforeEach(() => {
  vi.clearAllMocks()
  nextWebContentsId = 1
  docsQueryDirty.mockImplementation(() => Promise.resolve(false))
  requestDocsClose.mockImplementation(() => Promise.resolve(true))
  pdfIsDirty.mockImplementation(() => false)
  sheetsPendingEditCount.mockImplementation(() => 0)
  slidesIsDirty.mockImplementation(() => false)
  shellWindow = makeShellWindow()
  onChanged = vi.fn()
  applyMenuFor = vi.fn()
  manager = new TabManager(
    shellWindow as never,
    () => onChanged(),
    (kind) => applyMenuFor(kind),
  )
})

describe('initial state', () => {
  it('starts with only the non-closable, active Home tab', () => {
    expect(manager.list()).toEqual([
      { id: 'home', kind: 'home', title: 'Airy', closable: false, active: true },
    ])
  })
})

describe('opening tabs', () => {
  it('opens a docs tab, activates it, and attaches its view to the window', () => {
    const id = manager.openDocsTab()
    const tabs = manager.list()
    expect(tabs).toHaveLength(2)
    expect(tabs[1]).toMatchObject({
      id,
      kind: 'docs',
      title: 'Airy Docs',
      closable: true,
      active: true,
    })
    expect(tabs[0].active).toBe(false)
    expect(shellWindow.contentView.addChildView).toHaveBeenCalledTimes(1)
    expect(applyMenuFor).toHaveBeenLastCalledWith('docs')
    expect(onChanged).toHaveBeenCalled()
  })

  it('titles file-backed tabs with the file basename', () => {
    manager.openDocsTab('/tmp/report.docx')
    manager.openSheetsTab('/tmp/budget.xlsx')
    manager.openSlidesTab('/tmp/deck.pptx')
    manager.openPdfTab('/tmp/scan.pdf')
    expect(manager.list().map((t) => t.title)).toEqual([
      'Airy',
      'report.docx',
      'budget.xlsx',
      'deck.pptx',
      'scan.pdf',
    ])
  })

  it('uses module default titles for pathless tabs', () => {
    manager.openSheetsTab()
    manager.openSlidesTab()
    expect(manager.list().map((t) => t.title)).toEqual(['Airy', 'AI Sheets', 'AI Slides'])
  })

  it('assigns unique, monotonic tab ids', () => {
    const a = manager.openDocsTab()
    const b = manager.openSheetsTab()
    expect(a).not.toBe(b)
    expect(a).toBe('t1')
    expect(b).toBe('t2')
  })

  it('forwards the new-blank flag to the module', () => {
    manager.openDocsTab(undefined, { newBlank: true })
    expect(markDocsNewBlank).toHaveBeenCalledTimes(1)
    manager.openSheetsTab(undefined, { newBlank: true })
    expect(setSheetsNewBlank).toHaveBeenCalledTimes(1)
  })
})

describe('activation', () => {
  it('shows only the activated tab view and lays it out below the tab strip', () => {
    const docsId = manager.openDocsTab()
    const docsView = lastCreatedView(createDocsView)
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)

    manager.activateTab(docsId)
    expect(docsView.setVisible).toHaveBeenLastCalledWith(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(false)
    expect(docsView.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
    expect(manager.list().find((t) => t.id === docsId)?.active).toBe(true)
  })

  it('ignores activation of unknown tab ids', () => {
    onChanged.mockClear()
    manager.activateTab('nope')
    expect(onChanged).not.toHaveBeenCalled()
    expect(manager.list()[0].active).toBe(true)
  })

  it('routes the active webContents to the matching module', () => {
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    expect(setActiveSheetsWebContents).toHaveBeenLastCalledWith(sheetsView.webContents)

    manager.openSlidesTab()
    const slidesView = lastCreatedView(createSlidesView)
    expect(setActiveSlidesWebContents).toHaveBeenLastCalledWith(slidesView.webContents)
  })

  it('lets the active view cover the tab strip during HTML fullscreen', () => {
    manager.openSlidesTab()
    const view = lastCreatedView(createSlidesView)
    view.webContents.listeners.get('enter-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: 0,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT,
    })
    view.webContents.listeners.get('leave-html-full-screen')!()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })
  })
})

describe('window resize layout', () => {
  function resizeHandler(): () => void {
    const call = shellWindow.on.mock.calls.find((c) => c[0] === 'resize')
    expect(call).toBeDefined()
    return call![1] as () => void
  }

  it('re-lays out after resize bounds settle (Linux/X11 stale getContentBounds)', async () => {
    // On X11, `resize` fires before the WM applies maximize bounds, so the first
    // layout still sees the pre-maximize size. The deferred layout must pick up
    // the real size on the next turn (see issue #15).
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    view.setBounds.mockClear()

    let width = WINDOW_WIDTH
    let height = WINDOW_HEIGHT
    shellWindow.getContentBounds = () => ({ x: 0, y: 0, width, height })

    resizeHandler()()
    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: WINDOW_WIDTH,
      height: WINDOW_HEIGHT - TAB_STRIP_HEIGHT,
    })

    // Bounds update after the synchronous layout, as on X11 maximize.
    width = 1920
    height = 1080
    await new Promise<void>((resolve) => setImmediate(resolve))

    expect(view.setBounds).toHaveBeenLastCalledWith({
      x: 0,
      y: TAB_STRIP_HEIGHT,
      width: 1920,
      height: 1080 - TAB_STRIP_HEIGHT,
    })
    expect(view.setBounds).toHaveBeenCalledTimes(2)
  })

  it('skips deferred layout after the shell window is destroyed', async () => {
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    view.setBounds.mockClear()

    resizeHandler()()
    expect(view.setBounds).toHaveBeenCalledTimes(1)

    shellWindow.isDestroyed.mockReturnValue(true)
    await new Promise<void>((resolve) => setImmediate(resolve))
    expect(view.setBounds).toHaveBeenCalledTimes(1)
  })
})

describe('closing tabs', () => {
  it('never closes the Home tab', async () => {
    await manager.closeTab('home')
    expect(manager.list()).toHaveLength(1)

    manager.openHomeTab()
    manager.closeActiveTab()
    await Promise.resolve()
    expect(manager.list()).toHaveLength(1)
  })

  it('removes a clean tab and falls back to the previous tab', async () => {
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    const slidesId = manager.openSlidesTab()

    await manager.closeTab(slidesId)
    const tabs = manager.list()
    expect(tabs.map((t) => t.id)).toEqual(['home', 't1'])
    expect(tabs[1].active).toBe(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(true)
  })

  it('keeps the current tab active when closing a background tab', async () => {
    const sheetsId = manager.openSheetsTab()
    const slidesId = manager.openSlidesTab()
    await manager.closeTab(sheetsId)
    expect(manager.list().find((t) => t.id === slidesId)?.active).toBe(true)
  })

  it('detaches and destroys non-docs views on close', async () => {
    const id = manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    expect(view.webContents.close).toHaveBeenCalledTimes(1)
  })

  it('reclaims a closed markdown tab webContents outright (PERF-1657)', async () => {
    const id = manager.openMarkdownTab('/tmp/notes.md')
    const view = lastCreatedView(createMarkdownView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    // markdown tabs take the direct close path: the renderer process is
    // reclaimed deterministically, with no docs-style teardown navigation
    // leaving an orphaned webContents behind (BUG-409 / PERF-1640 do not apply)
    expect(view.webContents.close).toHaveBeenCalledTimes(1)
    expect(view.webContents.loadURL).not.toHaveBeenCalled()
  })

  it('reclaims a closed html tab webContents outright (PERF-1657)', async () => {
    const id = manager.openHtmlTab('/tmp/page.html')
    const view = lastCreatedView(createHtmlView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    // same non-docs path as markdown: close() is issued on the spot
    expect(view.webContents.close).toHaveBeenCalledTimes(1)
    expect(view.webContents.loadURL).not.toHaveBeenCalled()
  })

  it('detaches docs views without destroying the webContents synchronously (freeze workaround)', async () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    await manager.closeTab(id)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(view)
    // close() is never issued on the LIVE editor renderer — that is the path
    // that wedges Electron's UI thread (BUG-409); it may only run after the
    // renderer has been navigated away from the document
    expect(view.webContents.close).not.toHaveBeenCalled()
    // the orphaned renderer must be told to go inert (recovery-copy resurrection guard)
    expect(teardownDocsRenderer).toHaveBeenCalledWith(view.webContents)
  })

  it('navigates the orphaned docs renderer to about:blank, then reclaims the process (BUG-409 + PERF-1640)', async () => {
    const id = manager.openDocsTab('/tmp/thesis.docx')
    const view = lastCreatedView(createDocsView)
    await manager.closeTab(id)
    // the renderer is never destroyed while the document is live (freeze
    // workaround), so the whole docs app heap would live until quit; the
    // teardown navigation drops it
    expect(view.webContents.loadURL).toHaveBeenCalledWith('about:blank')
    // still inside the same tick: the wedged close path is not touched
    expect(view.webContents.close).not.toHaveBeenCalled()
    // once the navigation landed, the empty about:blank webContents is closed
    // so the orphaned Chromium renderer process is reclaimed (PERF-1640)
    await vi.waitFor(() => expect(view.webContents.close).toHaveBeenCalledTimes(1))
    // non-docs tabs are destroyed outright — no teardown navigation there
    const sheetsId = manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView)
    await manager.closeTab(sheetsId)
    expect(sheetsView.webContents.loadURL).not.toHaveBeenCalled()
    expect(sheetsView.webContents.close).toHaveBeenCalledTimes(1)
  })

  it('leaves the docs webContents orphaned when the teardown navigation never lands', async () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    // a crashed renderer cannot complete the navigation — closeTab must not
    // destroy the webContents behind a possibly-unanswered teardown
    view.webContents.loadURL.mockImplementation(() => Promise.reject(new Error('renderer gone')))
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {})
    try {
      await manager.closeTab(id)
      await new Promise((resolve) => setImmediate(resolve))
      await new Promise((resolve) => setImmediate(resolve))
      expect(teardownDocsRenderer).toHaveBeenCalledWith(view.webContents)
      expect(view.webContents.close).not.toHaveBeenCalled()
    } finally {
      errSpy.mockRestore()
    }
  })

  it('closes a clean docs tab after the async dirty query says clean', async () => {
    const id = manager.openDocsTab()
    await manager.closeTab(id)
    expect(docsQueryDirty).toHaveBeenCalledTimes(1)
    expect(requestDocsClose).not.toHaveBeenCalled()
    expect(manager.list()).toHaveLength(1)
  })

  it('keeps a dirty docs tab open when the user cancels the close guard', async () => {
    docsQueryDirty.mockImplementation(() => Promise.resolve(true))
    requestDocsClose.mockImplementation(() => Promise.resolve(false))
    const id = manager.openDocsTab()
    await manager.closeTab(id)
    expect(requestDocsClose).toHaveBeenCalledTimes(1)
    expect(manager.list().map((t) => t.id)).toEqual(['home', id])
  })

  it('activates a dirty background tab before showing its close guard', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    requestSheetsClose.mockImplementation(() => Promise.resolve(false))
    const sheetsId = manager.openSheetsTab()
    manager.openSlidesTab()

    await manager.closeTab(sheetsId)
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    // the guarded tab was brought into view for the prompt
    expect(manager.list().find((t) => t.id === sheetsId)?.active).toBe(true)
  })

  it('closes a dirty sheets tab when the guard resolves true', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    requestSheetsClose.mockImplementation(() => Promise.resolve(true))
    const id = manager.openSheetsTab()
    await manager.closeTab(id)
    expect(manager.list()).toHaveLength(1)
  })

  it('does not stack close guards while one prompt is pending', async () => {
    sheetsPendingEditCount.mockImplementation(() => 1)
    let resolveGuard!: (ok: boolean) => void
    requestSheetsClose.mockImplementation(
      () =>
        new Promise<boolean>((resolve) => {
          resolveGuard = resolve
        }),
    )
    const id = manager.openSheetsTab()
    const first = manager.closeTab(id)
    const second = manager.closeTab(id)
    resolveGuard(true)
    await Promise.all([first, second])
    expect(requestSheetsClose).toHaveBeenCalledTimes(1)
    expect(manager.list()).toHaveLength(1)
  })
})

describe('file path bookkeeping', () => {
  it('updates the tab title when a module opens a file in an existing tab', () => {
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    manager.setTabFileFor(view.webContents.id, '/tmp/final.docx')
    expect(manager.list()[1].title).toBe('final.docx')
    expect(manager.findDocsTabByPath('/tmp/final.docx')).toBe('t1')
  })

  it('ignores setTabFileFor for unknown webContents', () => {
    onChanged.mockClear()
    manager.setTabFileFor(999, '/tmp/x.docx')
    expect(onChanged).not.toHaveBeenCalled()
  })

  it('renames matching tabs and reports the affected views', () => {
    manager.openDocsTab('/tmp/old.docx')
    const view = lastCreatedView(createDocsView)
    const affected = manager.renameTabFile('/tmp/old.docx', '/tmp/new.docx')
    expect(affected).toEqual([{ kind: 'docs', webContents: view.webContents }])
    expect(manager.list()[1].title).toBe('new.docx')
    expect(manager.findDocsTabByPath('/tmp/new.docx')).toBe('t1')
    expect(manager.findDocsTabByPath('/tmp/old.docx')).toBeUndefined()
  })

  it('returns no affected views when nothing matches a rename', () => {
    onChanged.mockClear()
    expect(manager.renameTabFile('/tmp/none.docx', '/tmp/new.docx')).toEqual([])
    expect(onChanged).not.toHaveBeenCalled()
  })

  it('finds tabs by kind and path', () => {
    manager.openSheetsTab('/tmp/a.xlsx')
    manager.openSlidesTab('/tmp/b.pptx')
    manager.openPdfTab('/tmp/c.pdf')
    expect(manager.findSheetsTab()).toBe('t1')
    expect(manager.findSheetsTabByPath('/tmp/a.xlsx')).toBe('t1')
    expect(manager.findSlidesTabByPath('/tmp/b.pptx')).toBe('t2')
    expect(manager.findPdfTabByPath('/tmp/c.pdf')).toBe('t3')
    expect(manager.findPdfTabByPath('/tmp/missing.pdf')).toBeUndefined()
  })

  it('reloads an existing pdf tab so a re-export rereads the file from disk', () => {
    const id = manager.openPdfTab('/tmp/c.pdf')
    const view = lastCreatedView(createPdfView)
    manager.reloadTab(id)
    expect(clearPdfDirty).toHaveBeenCalledWith(view.webContents.id)
    expect(view.webContents.reload).toHaveBeenCalledTimes(1)
  })

  it('reports the active pdf tab with its id (so callers can re-activate it)', () => {
    const pdfId = manager.openPdfTab('/tmp/c.pdf')
    const active = manager.activePdfTab()
    expect(active?.id).toBe(pdfId)
    expect(active?.filePath).toBe('/tmp/c.pdf')
    manager.openDocsTab()
    expect(manager.activePdfTab()).toBeUndefined()
  })
})

describe('dirty-tab queries (shell close guard)', () => {
  it('lists only tabs whose module reports unsaved changes', () => {
    const dirtySheetsId = manager.openSheetsTab()
    const dirtyView = lastCreatedView(createSheetsView)
    manager.openSheetsTab()
    sheetsPendingEditCount.mockImplementation((id: number) =>
      id === dirtyView.webContents.id ? 3 : 0,
    )

    const dirty = manager.dirtySheetsTabs()
    expect(dirty).toEqual([{ id: dirtySheetsId, webContents: dirtyView.webContents }])
  })

  it('lists dirty pdf and slides tabs', () => {
    const pdfId = manager.openPdfTab('/tmp/c.pdf')
    const slidesId = manager.openSlidesTab()
    expect(manager.dirtyPdfTabs()).toEqual([])
    expect(manager.dirtySlidesTabs()).toEqual([])
    pdfIsDirty.mockImplementation(() => true)
    slidesIsDirty.mockImplementation(() => true)
    expect(manager.dirtyPdfTabs().map((t) => t.id)).toEqual([pdfId])
    expect(manager.dirtySlidesTabs().map((t) => t.id)).toEqual([slidesId])
  })

  it('lists every live docs tab for the async dirtiness sweep', () => {
    manager.openDocsTab()
    manager.openSheetsTab()
    manager.openDocsTab('/tmp/a.docx')
    expect(manager.docsTabs().map((t) => t.id)).toEqual(['t1', 't3'])
  })
})

describe('tab-switch accelerators', () => {
  const chord = (code: string, over: Record<string, unknown> = {}) =>
    ({
      type: 'keyDown',
      control: true,
      meta: false,
      alt: false,
      shift: false,
      code,
      ...over,
    }) as never

  function emitKey(view: FakeView, input: unknown): { preventDefault: () => void } {
    const handler = view.webContents.listeners.get('before-input-event')
    expect(handler).toBeDefined()
    const event = { preventDefault: vi.fn() }
    handler!(event, input)
    return event
  }

  it('attaches the before-input-event hook to every editor view', () => {
    manager.openDocsTab()
    manager.openSheetsTab()
    manager.openSlidesTab()
    for (const factory of [createDocsView, createSheetsView, createSlidesView]) {
      const view = lastCreatedView(factory)
      expect(view.webContents.on).toHaveBeenCalledWith('before-input-event', expect.any(Function))
    }
  })

  it('switches tabs from a keydown inside an editor view', () => {
    const docsId = manager.openDocsTab()
    manager.openSheetsTab()
    const sheetsView = lastCreatedView(createSheetsView) // active, owns keyboard focus

    const event = emitKey(sheetsView, chord('Digit2'))
    expect(event.preventDefault).toHaveBeenCalledTimes(1)
    expect(manager.list().find((t) => t.id === docsId)?.active).toBe(true)
  })

  it('keeps digits the active editor kind reserves', () => {
    manager.openHomeTab()
    manager.openDocsTab() // docs is active and reserves Ctrl+1
    const docsView = lastCreatedView(createDocsView)

    const event = emitKey(docsView, chord('Digit1'))
    expect(event.preventDefault).not.toHaveBeenCalled()
    expect(manager.list()[0].active).toBe(false)
  })

  it('detaches the hook when the tab closes (docs views outlive their tab)', async () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    await manager.closeTab(id)

    expect(view.webContents.removeListener).toHaveBeenCalledWith(
      'before-input-event',
      expect.any(Function),
    )
    expect(view.webContents.listeners.has('before-input-event')).toBe(false)
  })
})

describe('renderer crash recovery', () => {
  function makeCrashManager() {
    const onCrash = vi.fn()
    const errorPageBody = vi.fn(() => 'This tab stopped unexpectedly.')
    const manager = new TabManager(
      shellWindow as never,
      () => onChanged(),
      (kind) => applyMenuFor(kind),
      undefined,
      { errorPageBody, onCrash },
    )
    return { manager, onCrash, errorPageBody }
  }

  function emitCrash(view: FakeView, reason: string): void {
    const handler = view.webContents.listeners.get('render-process-gone')
    expect(handler).toBeDefined()
    handler!({}, { reason })
  }

  it('marks a crashed tab, shows the error page and notifies the shell', () => {
    const { manager, onCrash } = makeCrashManager()
    const id = manager.openDocsTab('/tmp/report.docx')
    const view = lastCreatedView(createDocsView)
    expect(manager.isTabCrashed(id)).toBe(false)

    emitCrash(view, 'oom')

    expect(manager.isTabCrashed(id)).toBe(true)
    expect(onCrash).toHaveBeenCalledWith({
      id,
      kind: 'docs',
      title: 'report.docx',
      reason: 'oom',
    })
    // in-tab error state replaces the dead renderer content
    expect(view.webContents.loadURL).toHaveBeenCalledTimes(1)
    const url = view.webContents.loadURL.mock.calls[0]![0] as string
    expect(url.startsWith('data:text/html;charset=utf-8,')).toBe(true)
    expect(decodeURIComponent(url)).toContain('This tab stopped unexpectedly.')
  })

  it('reloadTab clears the crashed state and restarts the renderer', () => {
    const { manager } = makeCrashManager()
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    emitCrash(view, 'crashed')

    manager.reloadTab(id)

    expect(manager.isTabCrashed(id)).toBe(false)
    expect(view.webContents.reload).toHaveBeenCalledTimes(1)
  })

  it('ignores intentional teardown reasons', () => {
    const { manager, onCrash } = makeCrashManager()
    manager.openDocsTab()
    const view = lastCreatedView(createDocsView)

    emitCrash(view, 'clean-exit')

    expect(onCrash).not.toHaveBeenCalled()
    expect(view.webContents.loadURL).not.toHaveBeenCalled()
  })

  it('prompts only once for a repeated crash signal', () => {
    const { manager, onCrash } = makeCrashManager()
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    emitCrash(view, 'oom')
    emitCrash(view, 'oom')

    expect(onCrash).toHaveBeenCalledTimes(1)
    expect(manager.isTabCrashed(id)).toBe(true)
  })
})

describe('moving a tab to another window (detach/adopt)', () => {
  it('detachTab lifts the tab with its live view and falls back to the previous tab', () => {
    const sheetsId = manager.openSheetsTab('/tmp/a.xlsx')
    const sheetsView = lastCreatedView(createSheetsView)
    const slidesId = manager.openSlidesTab('/tmp/b.pptx')
    const slidesView = lastCreatedView(createSlidesView)

    const detached = manager.detachTab(slidesId)!

    expect(detached).toMatchObject({ kind: 'slides', title: 'b.pptx', filePath: '/tmp/b.pptx' })
    expect(detached.view).toBe(slidesView)
    // removed from the source strip, previous tab re-activated (and shown)
    expect(manager.list().map((t) => t.id)).toEqual(['home', sheetsId])
    expect(manager.list().find((t) => t.id === sheetsId)?.active).toBe(true)
    expect(sheetsView.setVisible).toHaveBeenLastCalledWith(true)
    // the view left this window's content view (hidden, renderer still alive)
    expect(shellWindow.contentView.removeChildView).toHaveBeenCalledWith(slidesView)
    expect(slidesView.webContents.close).not.toHaveBeenCalled()
    expect(slidesView.setVisible).toHaveBeenLastCalledWith(false)
  })

  it('never detaches the Home tab and reports unknown ids as null', () => {
    expect(manager.detachTab('home')).toBeNull()
    expect(manager.detachTab('nope')).toBeNull()
  })

  it('detach does not fire onTabClosed (the staged file survives the move)', () => {
    const onTabClosed = vi.fn()
    manager.onTabClosed = onTabClosed
    const id = manager.openSheetsTab('/tmp/a.xlsx')
    manager.detachTab(id)
    expect(onTabClosed).not.toHaveBeenCalled()
  })

  it('detach detaches the accelerator hook so the adopter can re-attach its own', () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    manager.detachTab(id)
    expect(view.webContents.listeners.has('before-input-event')).toBe(false)
  })

  it('adoptTab re-parents the view into the adopting window and activates it', () => {
    const detached = (() => {
      const id = manager.openSheetsTab('/tmp/a.xlsx')
      return manager.detachTab(id)!
    })()

    const window2 = makeShellWindow()
    const manager2 = new TabManager(
      window2 as never,
      () => onChanged(),
      (kind) => applyMenuFor(kind),
    )
    const adoptedId = manager2.adoptTab(detached)

    expect(manager2.list()).toEqual([
      { id: 'home', kind: 'home', title: 'Airy', closable: false, active: false },
      { id: adoptedId, kind: 'sheets', title: 'a.xlsx', closable: true, active: true },
    ])
    expect(window2.contentView.addChildView).toHaveBeenCalledWith(detached.view)
    expect(detached.view.setVisible).toHaveBeenLastCalledWith(true)
    expect(detached.view.webContents.on).toHaveBeenCalledWith(
      'before-input-event',
      expect.any(Function),
    )
  })

  it('a moved tab switches tabs in its NEW window only', async () => {
    const docsId = manager.openDocsTab()
    manager.openSheetsTab('/tmp/a.xlsx')
    const sheetsView = lastCreatedView(createSheetsView)
    const detached = manager.detachTab(manager.list().find((t) => t.kind === 'sheets')!.id)!

    const window2 = makeShellWindow()
    const manager2 = new TabManager(
      window2 as never,
      () => {},
      (kind) => applyMenuFor(kind),
    )
    manager2.openDocsTab('/tmp/z.docx')
    manager2.adoptTab(detached)

    // Ctrl+2 inside the moved view: window 2 has [home, z.docx, a.xlsx]
    const handler = sheetsView.webContents.listeners.get('before-input-event')!
    const event = { preventDefault: vi.fn() }
    handler(event, {
      type: 'keyDown',
      control: true,
      meta: false,
      alt: false,
      shift: false,
      code: 'Digit2',
    })
    expect(event.preventDefault).toHaveBeenCalled()
    expect(manager2.list().find((t) => t.title === 'z.docx')?.active).toBe(true)
    expect(manager.list().find((t) => t.id === docsId)?.active).toBe(true) // source untouched
  })
})

describe('view watcher teardown across window moves (BUG-1106)', () => {
  /** net listeners still registered for an event: on() calls minus matching
   *  removeListener() calls — the fake Map alone would hide stacked pairs */
  function listenerCount(view: FakeView, event: string): number {
    const on = view.webContents.on.mock.calls.filter(([ev]) => ev === event).length
    const off = view.webContents.removeListener.mock.calls.filter(([ev]) => ev === event).length
    return on - off
  }

  const WATCHED_EVENTS = [
    'enter-html-full-screen',
    'leave-html-full-screen',
    'render-process-gone',
    'before-input-event',
  ]

  function makeManagerWithWindow() {
    const window = makeShellWindow()
    const mgr = new TabManager(
      window as never,
      () => {},
      (kind) => applyMenuFor(kind),
    )
    return { window, mgr }
  }

  it('repeated moves keep exactly one watcher set on the view', () => {
    const a = makeManagerWithWindow()
    const b = makeManagerWithWindow()
    const id = a.mgr.openSlidesTab('/tmp/deck.pptx')
    const view = lastCreatedView(createSlidesView)

    // three round trips A → B → A → B
    let current = a
    let currentId = id
    for (let i = 0; i < 3; i += 1) {
      const detached = current.mgr.detachTab(currentId)!
      current = current === a ? b : a
      currentId = current.mgr.adoptTab(detached)
    }

    for (const event of WATCHED_EVENTS) {
      expect(listenerCount(view, event), event).toBe(1)
    }
  })

  it('detach leaves no watchers behind on the source manager', () => {
    const id = manager.openSlidesTab('/tmp/deck.pptx')
    const view = lastCreatedView(createSlidesView)
    manager.detachTab(id)
    for (const event of WATCHED_EVENTS) {
      expect(listenerCount(view, event), event).toBe(0)
    }
  })

  it('only the adopting manager reacts after a move (the old pair is gone)', () => {
    const onCrashA = vi.fn()
    const onCrashB = vi.fn()
    const errorPage = () => 'err'
    const windowA = makeShellWindow()
    const managerA = new TabManager(
      windowA as never,
      () => {},
      (kind) => applyMenuFor(kind),
      undefined,
      { errorPageBody: errorPage, onCrash: onCrashA },
    )
    const windowB = makeShellWindow()
    const managerB = new TabManager(
      windowB as never,
      () => {},
      (kind) => applyMenuFor(kind),
      undefined,
      { errorPageBody: errorPage, onCrash: onCrashB },
    )

    const idA = managerA.openDocsTab('/tmp/report.docx')
    const view = lastCreatedView(createDocsView)
    const detached = managerA.detachTab(idA)!
    const idB = managerB.adoptTab(detached)

    view.webContents.listeners.get('render-process-gone')!({}, { reason: 'oom' })

    expect(onCrashB).toHaveBeenCalledWith({
      id: idB,
      kind: 'docs',
      title: 'report.docx',
      reason: 'oom',
    })
    expect(onCrashA).not.toHaveBeenCalled()
    expect(managerA.isTabCrashed(idA)).toBe(false)
    expect(managerB.isTabCrashed(idB)).toBe(true)
  })

  it('closing a torn-down docs tab removes its watchers too', async () => {
    const id = manager.openDocsTab()
    const view = lastCreatedView(createDocsView)
    await manager.closeTab(id)
    for (const event of WATCHED_EVENTS) {
      expect(listenerCount(view, event), event).toBe(0)
    }
  })
})

describe('move-tab chord wiring', () => {
  const chord = {
    type: 'keyDown',
    control: true,
    meta: false,
    alt: false,
    shift: true,
    code: 'KeyK',
  }

  function emitKey(view: FakeView, input: unknown): { preventDefault: () => void } {
    const handler = view.webContents.listeners.get('before-input-event')
    expect(handler).toBeDefined()
    const event = { preventDefault: vi.fn() }
    handler!(event, input)
    return event
  }

  it('fires onMoveTabToNewWindow for the active editor tab and eats the chord', () => {
    const onMove = vi.fn()
    manager.onMoveTabToNewWindow = onMove
    const id = manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)

    const event = emitKey(view, chord)
    expect(event.preventDefault).toHaveBeenCalledTimes(1)
    expect(onMove).toHaveBeenCalledWith(id)
  })

  it('eats the chord but stays cold while Home is the active tab', () => {
    const onMove = vi.fn()
    manager.onMoveTabToNewWindow = onMove
    manager.openSheetsTab()
    const view = lastCreatedView(createSheetsView)
    manager.activateTab('home')

    const event = emitKey(view, chord)
    expect(event.preventDefault).toHaveBeenCalledTimes(1)
    expect(onMove).not.toHaveBeenCalled()
  })
})

import { basename } from 'node:path'
import { BrowserWindow } from 'electron'
import type {
  Event as ElectronEvent,
  Input,
  Rectangle,
  WebContents,
  WebContentsView,
} from 'electron'
import {
  crashErrorPageUrl,
  grantRendererFileAccess,
  isRecoverableRendererCrash,
  voidLoad,
} from '@airy-office/electron-utils'

import {
  createDocsView,
  docsQueryDirty,
  markDocsNewBlank,
  queueDocsAiContent,
  recordRecentFile,
  requestDocsClose,
  setActiveDocsResolver,
  teardownDocsRenderer,
} from '../../../docs/src/main/docs-main'
import type { AiDocContent } from '../../../docs/src/shared/ipc'
import {
  createMarkdownView,
  markdownIsDirty,
  requestMarkdownClose,
} from '../../../markdown/src/main/markdown-main'
import {
  createHtmlPresentView,
  createHtmlView,
  htmlIsDirty,
  requestHtmlClose,
} from '../../../html/src/main/html-main'
import {
  createPdfView,
  clearPdfDirty,
  pdfIsDirty,
  requestPdfClose,
} from '../../../pdf/src/main/pdf-main'
import {
  createSheetsView,
  queueWorkbookForView,
  requestSheetsClose,
  setActiveSheetsWebContents,
  setSheetsNewBlank,
  sheetsPendingEditCount,
} from '../../../sheets/src/main/sheets-main'
import {
  createSlidesView,
  requestSlidesClose,
  setActiveSlidesWebContents,
  slidesIsDirty,
} from '../../../slides/src/main/slides-main'
import { tabSwitchTargetForInput } from './tab-accelerators'
import type { TabKind, TabSummary } from '../shared/tabs-api'

interface TabRecord {
  id: string
  kind: TabKind
  /** null for the Home tab — it's rendered by the shell window's own webContents */
  view: WebContentsView | null
  title: string
  filePath?: string
  /** chrome-free Present tab: no file, no editor menu or save/export targets */
  present?: boolean
  /** renderer crashed (oom/crashed): awaiting the user's Reload/Close decision */
  crashed?: boolean
}

/** Renderer-crash recovery hooks, supplied by the shell (localized prompt + error page). */
export interface TabCrashUi {
  /** body text for the in-tab error page shown while awaiting the decision */
  errorPageBody: () => string
  /** notified once per crash; the shell shows the Reload/Close dialog */
  onCrash: (info: { id: string; kind: TabKind; title: string; reason: string }) => void
}

/** must match the tab strip's rendered height (apps/shell/src/renderer/src/TabBar.tsx) */
const TAB_STRIP_HEIGHT = 40
const HOME_ID = 'home'

/**
 * Owns every open tab (Home + docs + sheets) inside the shell's single
 * BrowserWindow. Docs/sheets tabs are WebContentsView children of that
 * window; only the active one is visible at a time. Home has no view of its
 * own — hiding every other tab reveals the shell window's own content.
 */
export class TabManager {
  private readonly tabs: TabRecord[] = [{ id: HOME_ID, kind: 'home', view: null, title: 'Airy' }]
  private activeId: string = HOME_ID
  private nextId = 1
  /** tab whose page entered HTML fullscreen (e.g. slides slideshow) — its view covers the tab strip */
  private htmlFullScreenId: string | null = null
  /** webContents ids whose view must cover the tab strip without HTML fullscreen
   *  (slides show: the window snaps via simpleFullScreen and asks for the bleed
   *  over IPC, since requestFullscreen would animate the native transition) */
  private readonly bleedWcIds = new Set<number>()
  /** tabs mid unsaved-changes prompt, so a second close click doesn't stack dialogs */
  private readonly closingIds = new Set<string>()
  /** live before-input-event handlers per view webContents id (detach on close) */
  private readonly acceleratorHandlers = new Map<
    number,
    (event: ElectronEvent, input: Input) => void
  >()

  constructor(
    private readonly shellWindow: BrowserWindow,
    private readonly onChanged: () => void,
    private readonly applyMenuFor: (kind: TabKind) => void,
    /** localized placeholder title for a tab that has no file yet */
    private readonly untitledTitleFor?: (kind: TabKind) => string,
    /** renderer-crash recovery (error page + Reload/Close prompt), shell-provided */
    private readonly crashUi?: TabCrashUi,
  ) {
    // Layout once synchronously for macOS/Windows (bounds are already correct),
    // then once more on the next tick. On Linux/X11, `resize` fires before the
    // window manager applies the new size, so getContentBounds() is still the
    // pre-maximize size inside the handler and a follow-up layout is required.
    // See https://github.com/genspark-ai/genoffice/issues/15
    shellWindow.on('resize', () => {
      this.layout()
      setImmediate(() => this.layout())
    })
  }

  private untitled(kind: TabKind, fallback: string): string {
    return this.untitledTitleFor?.(kind) ?? fallback
  }

  private contentBounds(): Rectangle {
    const { width, height } = this.shellWindow.getContentBounds()
    if (this.htmlFullScreenId !== null && this.htmlFullScreenId === this.activeId) {
      return { x: 0, y: 0, width, height }
    }
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view && this.bleedWcIds.has(active.view.webContents.id)) {
      return { x: 0, y: 0, width, height }
    }
    return { x: 0, y: TAB_STRIP_HEIGHT, width, height: Math.max(0, height - TAB_STRIP_HEIGHT) }
  }

  /** Grow/restore a tab view over the tab strip on request (slides show fullscreen) */
  setContentBleed(wc: WebContents, on: boolean): void {
    if (on) this.bleedWcIds.add(wc.id)
    else this.bleedWcIds.delete(wc.id)
    this.layout()
  }

  /**
   * When a tab's page enters HTML fullscreen (the slides slideshow calls requestFullscreen),
   * grow its view over the tab strip so nothing of the shell chrome shows;
   * restore the normal bounds on leave.
   */
  private trackHtmlFullScreen(id: string, view: WebContentsView): void {
    view.webContents.on('enter-html-full-screen', () => {
      this.htmlFullScreenId = id
      this.layout()
    })
    view.webContents.on('leave-html-full-screen', () => {
      if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
      this.layout()
    })
  }

  /**
   * Crash recovery: a renderer lost to oom/crashed would otherwise stay as a
   * blank zombie tab. Replace its content with an error page and hand the
   * decision to the shell (Reload restarts the renderer, Close drops the
   * tab). Intentional teardown reasons never prompt.
   */
  private watchRendererCrash(id: string, view: WebContentsView): void {
    view.webContents.on('render-process-gone', (_event, details) => {
      if (!isRecoverableRendererCrash(details.reason)) return
      const tab = this.tabs.find((t) => t.id === id)
      if (!tab || tab.crashed) return
      tab.crashed = true
      if (this.htmlFullScreenId === id) {
        this.htmlFullScreenId = null
        this.layout()
      }
      voidLoad(
        view.webContents.loadURL(crashErrorPageUrl(this.crashUi?.errorPageBody() ?? '')),
        `crash error page for tab ${id}`,
      )
      this.crashUi?.onCrash({ id, kind: tab.kind, title: tab.title, reason: details.reason })
    })
  }

  /** whether a tab's renderer crashed and is awaiting recovery (observability for tests) */
  isTabCrashed(id: string): boolean {
    return this.tabs.find((t) => t.id === id)?.crashed === true
  }

  /**
   * Ctrl/Cmd+1..8 → tab N, Ctrl/Cmd+9 → last tab, inside THIS view. Editor
   * tabs are sibling WebContentsViews with their own webContents — when one
   * has keyboard focus, keydowns never reach the before-input-event hook the
   * shell registers on its own (Home) webContents, and the editor-built menus
   * carry no digit accelerators. So every view gets the same hook here, at
   * creation, sharing the pure decision in tab-accelerators.ts (reserved
   * digits of the ACTIVE tab's kind stay with the editor). Detached on close
   * (docs views outlive their tab). Standalone editor windows are not hooked:
   * they are owned by the editor mains, own no tab strip, and the shell's tab
   * manager cannot be reached from them without cross-module plumbing for a
   * switch the user could not see.
   */
  private watchTabAccelerators(view: WebContentsView): void {
    const handler = (event: ElectronEvent, input: Input): void => {
      const target = tabSwitchTargetForInput(input, this.list())
      if (target === null) return
      event.preventDefault()
      this.activateTab(target)
    }
    this.acceleratorHandlers.set(view.webContents.id, handler)
    view.webContents.on('before-input-event', handler)
  }

  /** drop a closed view's accelerator hook (webContents may outlive the tab) */
  private detachTabAccelerators(wc: WebContents): void {
    // webContents.close() already dropped everything; this matters for the
    // docs teardown path, which detaches the view without destroying it
    const handler = this.acceleratorHandlers.get(wc.id)
    if (!handler) return
    this.acceleratorHandlers.delete(wc.id)
    if (!wc.isDestroyed()) wc.removeListener('before-input-event', handler)
  }

  /** re-fit the active tab's view after a window resize */
  layout(): void {
    // Deferred resize layouts can land after the shell window was closed.
    if (this.shellWindow.isDestroyed()) return
    const active = this.tabs.find((t) => t.id === this.activeId)
    if (active?.view) active.view.setBounds(this.contentBounds())
  }

  list(): TabSummary[] {
    return this.tabs.map((t) => ({
      id: t.id,
      kind: t.kind,
      title: t.title,
      closable: t.id !== HOME_ID,
      active: t.id === this.activeId,
    }))
  }

  /** id of the active tab (session persistence) */
  activeTabId(): string {
    return this.activeId
  }

  /** current backing file of the tab owning this webContents (staged-save rebind checks) */
  tabFilePathFor(webContentsId: number): string | undefined {
    return this.tabs.find((t) => t.view?.webContents.id === webContentsId)?.filePath
  }

  /** full record snapshot for shell-side menus (tab context-menu enablement) */
  tabInfo(id: string): { id: string; kind: TabKind; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === id)
    return tab ? { id: tab.id, kind: tab.kind, filePath: tab.filePath } : undefined
  }

  /**
   * Open the same backing file in a new tab (context-menu Duplicate).
   * Untitled/in-memory and present tabs have no file — returns false.
   */
  duplicateTab(id: string): boolean {
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab?.filePath || tab.present) return false
    // same grants/recents a routed open would apply — but bypassing the
    // open-by-path dedupe, which would just activate this tab
    grantRendererFileAccess(tab.filePath)
    recordRecentFile(tab.filePath)
    switch (tab.kind) {
      case 'docs':
        this.openDocsTab(tab.filePath)
        break
      case 'sheets':
        this.openSheetsTab(tab.filePath)
        break
      case 'slides':
        this.openSlidesTab(tab.filePath)
        break
      case 'pdf':
        this.openPdfTab(tab.filePath)
        break
      case 'markdown':
        this.openMarkdownTab(tab.filePath)
        break
      case 'html':
        this.openHtmlTab(tab.filePath)
        break
      default:
        return false
    }
    return true
  }

  /** whether any tab currently shows this file (staged-survivor purge at launch) */
  hasTabForPath(path: string): boolean {
    return this.tabs.some((t) => t.filePath === path)
  }

  /**
   * Fired after a tab was removed (per-tab close — a whole-window quit closes
   * the window instead, so this does not fire there). The shell uses it to
   * delete staged untitled files when their tab closes without a first save.
   */
  onTabClosed?: (tab: { id: string; kind: TabKind; filePath?: string }) => void

  /** file-backed tabs in strip order (session persistence; untitled/present tabs have no file) */
  sessionTabs(): Array<{ id: string; kind: TabKind; filePath: string | undefined }> {
    return this.tabs.map((t) => ({ id: t.id, kind: t.kind, filePath: t.filePath }))
  }

  /** the tab showing this file for a given kind, if any (session restore activation) */
  findTabIdByPath(kind: TabKind, path: string): string | undefined {
    return this.tabs.find((t) => t.kind === kind && t.filePath === path)?.id
  }

  openHomeTab(): void {
    this.activateTab(HOME_ID)
  }

  openDocsTab(
    openPath?: string,
    options?: { newBlank?: boolean; aiContent?: AiDocContent },
  ): string {
    const view = createDocsView(openPath)
    const id = `t${this.nextId++}`
    if (options?.newBlank) markDocsNewBlank(view.webContents.id)
    if (options?.aiContent) queueDocsAiContent(view.webContents.id, options.aiContent)
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({
      id,
      kind: 'docs',
      view,
      title: openPath ? basename(openPath) : this.untitled('docs', 'Airy Docs'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openSheetsTab(openPath?: string, options?: { newBlank?: boolean }): string {
    if (options?.newBlank) setSheetsNewBlank()
    const view = createSheetsView({ includeAiHandlers: false })
    // bind the path to this tab's webContents: a multi-select Open creates
    // several sheets tabs in one loop, so a single global path would be
    // overwritten before the earlier tabs consume it
    if (openPath) queueWorkbookForView(view.webContents, openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({
      id,
      kind: 'sheets',
      view,
      title: openPath ? basename(openPath) : this.untitled('sheets', 'AI Sheets'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openSlidesTab(openPath?: string): string {
    const view = createSlidesView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({
      id,
      kind: 'slides',
      view,
      title: openPath ? basename(openPath) : this.untitled('slides', 'AI Slides'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openPdfTab(openPath: string): string {
    const view = createPdfView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({ id, kind: 'pdf', view, title: basename(openPath), filePath: openPath })
    this.activateTab(id)
    return id
  }

  /** Remount the tab's renderer so it re-reads its file from disk (View > Reload;
   *  also the crash-recovery "Reload" action — clears the crashed state). */
  reloadTab(id: string): void {
    const tab = this.tabs.find((t) => t.id === id)
    const wc = tab?.view?.webContents
    if (!wc || wc.isDestroyed()) return
    if (tab.kind === 'pdf') clearPdfDirty(wc.id)
    tab.crashed = false
    wc.reload()
  }

  openMarkdownTab(openPath?: string): string {
    const view = createMarkdownView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({
      id,
      kind: 'markdown',
      view,
      title: openPath ? basename(openPath) : this.untitled('markdown', 'AI Markdown'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  openHtmlTab(openPath?: string): string {
    const view = createHtmlView(openPath)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({
      id,
      kind: 'html',
      view,
      title: openPath ? basename(openPath) : this.untitled('html', 'AI HTML'),
      filePath: openPath,
    })
    this.activateTab(id)
    return id
  }

  /** Present → New tab: a chrome-free html tab showing the owner tab's live preview */
  openHtmlPresentTab(owner: WebContents, title: string): string {
    const view = createHtmlPresentView(owner, title)
    const id = `t${this.nextId++}`
    this.shellWindow.contentView.addChildView(view)
    view.setVisible(false)
    this.trackHtmlFullScreen(id, view)
    this.watchRendererCrash(id, view)
    this.watchTabAccelerators(view)
    this.tabs.push({
      id,
      kind: 'html',
      view,
      title: title || this.untitled('html', 'AI HTML'),
      present: true,
    })
    this.activateTab(id)
    return id
  }

  activateTab(id: string): void {
    const target = this.tabs.find((t) => t.id === id)
    if (!target) return
    for (const t of this.tabs) t.view?.setVisible(t.id === id)
    if (target.view) target.view.setBounds(this.contentBounds())
    this.activeId = id
    this.refreshActiveTargets()
    this.onChanged()
  }

  /** Re-point the process-global active-editor targets and the app menu at this
   *  window's active tab. Called on every activation and on shell-window focus:
   *  a detached editor window ("Open in New Window") claims the same globals
   *  while it is focused. */
  refreshActiveTargets(): void {
    const target = this.tabs.find((t) => t.id === this.activeId)
    if (!target) return
    setActiveDocsResolver(target.kind === 'docs' ? () => target.view!.webContents : () => null)
    if (target.kind === 'sheets' && target.view) setActiveSheetsWebContents(target.view.webContents)
    if (target.kind === 'slides' && target.view) setActiveSlidesWebContents(target.view.webContents)
    this.applyMenuFor(target.present ? 'home' : target.kind)
  }

  /** move a tab to a new index in the strip; Home is pinned at index 0 */
  reorderTab(id: string, toIndex: number): void {
    if (id === HOME_ID) return
    const fromIndex = this.tabs.findIndex((t) => t.id === id)
    if (fromIndex < 0) return
    const clamped = Math.min(Math.max(Math.trunc(toIndex), 1), this.tabs.length - 1)
    if (clamped === fromIndex) return
    const [moved] = this.tabs.splice(fromIndex, 1)
    this.tabs.splice(clamped, 0, moved)
    this.onChanged()
  }

  tabIdForWebContents(webContentsId: number): string | undefined {
    return this.tabs.find((t) => t.view?.webContents.id === webContentsId)?.id
  }

  /** a module opened a file inside an existing tab (⌘O / queued path) — sync title + dedupe path */
  setTabFileFor(webContentsId: number, filePath: string): void {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab) return
    tab.filePath = filePath
    tab.title = basename(filePath)
    this.onChanged()
  }

  /** an untitled document named itself before its first save (html: from the first AI request) */
  setTabTitleFor(webContentsId: number, title: string): void {
    const tab = this.tabs.find((t) => t.view?.webContents.id === webContentsId)
    if (!tab || tab.filePath || tab.title === title) return
    tab.title = title
    this.onChanged()
  }

  /** a file was renamed on disk (rename from the Home list) — sync any open tab's title/path;
   *  returns the affected views so callers can notify the embedded editors */
  renameTabFile(
    oldPath: string,
    newPath: string,
  ): Array<{ kind: TabKind; webContents: WebContents }> {
    const affected: Array<{ kind: TabKind; webContents: WebContents }> = []
    for (const tab of this.tabs) {
      if (tab.filePath !== oldPath) continue
      tab.filePath = newPath
      tab.title = basename(newPath)
      if (tab.view) affected.push({ kind: tab.kind, webContents: tab.view.webContents })
    }
    if (affected.length > 0) this.onChanged()
    return affected
  }

  /** sheets tabs whose renderer reports unsaved journal edits (shell-close guard) */
  dirtySheetsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter(
        (t) => t.kind === 'sheets' && t.view && sheetsPendingEditCount(t.view.webContents.id) > 0,
      )
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** pdf tabs whose renderer reports unsaved markups/form edits (shell-close guard) */
  dirtyPdfTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'pdf' && t.view && pdfIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** markdown tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyMarkdownTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'markdown' && t.view && markdownIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** html tabs whose renderer reports unsaved edits (shell-close guard) */
  dirtyHtmlTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'html' && t.view && htmlIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** slides tabs whose main-process session has unsaved edits (shell-close guard) */
  dirtySlidesTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'slides' && t.view && slidesIsDirty(t.view.webContents.id))
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** all live docs tabs — dirtiness lives renderer-side, caller queries async (shell-close guard) */
  docsTabs(): Array<{ id: string; webContents: WebContents }> {
    return this.tabs
      .filter((t) => t.kind === 'docs' && t.view)
      .map((t) => ({ id: t.id, webContents: t.view!.webContents }))
  }

  /** open docs tabs with their files, for the external bridge's list method */
  docsTabSummaries(): Array<{
    id: string
    title: string
    filePath: string | null
    active: boolean
  }> {
    return this.tabs
      .filter((t) => t.kind === 'docs' && t.view)
      .map((t) => ({
        id: t.id,
        title: t.title,
        filePath: t.filePath ?? null,
        active: t.id === this.activeId,
      }))
  }

  /** the active tab's docs view, when the active tab is a docs document (live-bridge target) */
  activeDocsTab(): { id: string; webContents: WebContents } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'docs' && tab.view
      ? { id: tab.id, webContents: tab.view.webContents }
      : undefined
  }

  /** closes whichever tab is currently active; no-op for Home (Cmd+W target) */
  closeActiveTab(): void {
    void this.closeTab(this.activeId)
  }

  async closeTab(id: string): Promise<void> {
    if (id === HOME_ID) return
    const tab = this.tabs.find((t) => t.id === id)
    if (!tab || this.closingIds.has(id)) return
    let closeGuard =
      tab.view &&
      (tab.kind === 'sheets' && sheetsPendingEditCount(tab.view.webContents.id) > 0
        ? requestSheetsClose
        : tab.kind === 'pdf' && pdfIsDirty(tab.view.webContents.id)
          ? requestPdfClose
          : tab.kind === 'markdown' && markdownIsDirty(tab.view.webContents.id)
            ? requestMarkdownClose
            : tab.kind === 'html' && htmlIsDirty(tab.view.webContents.id)
              ? requestHtmlClose
              : tab.kind === 'slides' && slidesIsDirty(tab.view.webContents.id)
                ? requestSlidesClose
                : null)
    // docs dirty state lives in the renderer and needs an async query; skip the guard when clean (avoids a flash activation)
    if (!closeGuard && tab.kind === 'docs' && tab.view) {
      this.closingIds.add(id)
      try {
        if (await docsQueryDirty(tab.view.webContents)) closeGuard = requestDocsClose
      } finally {
        this.closingIds.delete(id)
      }
    }
    if (closeGuard && tab.view) {
      // Bring the tab into view so the save prompt has visible context.
      if (this.activeId !== id) this.activateTab(id)
      this.closingIds.add(id)
      try {
        if (!(await closeGuard(tab.view.webContents, this.shellWindow))) return
      } finally {
        this.closingIds.delete(id)
      }
    }
    const idx = this.tabs.findIndex((t) => t.id === id)
    if (idx < 0) return
    if (this.htmlFullScreenId === id) this.htmlFullScreenId = null
    const [removed] = this.tabs.splice(idx, 1)
    this.onTabClosed?.({ id: removed.id, kind: removed.kind, filePath: removed.filePath })
    if (this.activeId === id) {
      const fallback = this.tabs[idx - 1] ?? this.tabs[0]
      this.activateTab(fallback.id)
    } else {
      this.onChanged()
    }
    if (removed.view) {
      removed.view.setVisible(false)
      this.shellWindow.contentView.removeChildView(removed.view)
      this.detachTabAccelerators(removed.view.webContents)
      if (removed.kind === 'docs') {
        // webContents.close()/.destroy() on a closed docs tab wedges Electron's whole
        // UI thread in a native modal run loop (reproduced consistently; survives
        // close() vs destroy(), teardown ordering, deferring, and disabling
        // accessibility support — looks like an upstream WebContentsView/Chromium
        // issue, not something fixable from here). Detaching without destroying
        // avoids the freeze; the orphaned webContents is reclaimed when the app quits.
        teardownDocsRenderer(removed.view.webContents)
      } else {
        removed.view.webContents.close()
      }
    }
  }

  findDocsTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'docs' && t.filePath === path)?.id
  }

  findSheetsTab(): string | undefined {
    return this.tabs.find((t) => t.kind === 'sheets')?.id
  }

  findSheetsTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'sheets' && t.filePath === path)?.id
  }

  findSlidesTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'slides' && t.filePath === path)?.id
  }

  findPdfTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'pdf' && t.filePath === path)?.id
  }

  findMarkdownTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'markdown' && t.filePath === path)?.id
  }

  findHtmlTabByPath(path: string): string | undefined {
    return this.tabs.find((t) => t.kind === 'html' && t.filePath === path)?.id
  }

  /** the active tab's html view, if the active tab is html (html menu target) */
  activeHtmlTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'html' && tab.view && !tab.present
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }

  /** the active tab's markdown view, if the active tab is markdown (markdown menu target) */
  activeMarkdownTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'markdown' && tab.view
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }

  /** the active tab's pdf view, if the active tab is a pdf (pdf menu target) */
  activePdfTab(): { id: string; webContents: WebContents; filePath?: string } | undefined {
    const tab = this.tabs.find((t) => t.id === this.activeId)
    return tab?.kind === 'pdf' && tab.view
      ? { id: tab.id, webContents: tab.view.webContents, filePath: tab.filePath }
      : undefined
  }
}

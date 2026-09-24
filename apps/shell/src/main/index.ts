import { execSync, spawn } from 'node:child_process'
import {
  copyFileSync,
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  writeFileSync,
} from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'
import {
  BrowserWindow,
  Menu,
  app,
  dialog,
  ipcMain,
  nativeImage,
  nativeTheme,
  screen,
  session,
  shell,
  webContents,
} from 'electron'
import type { MenuItemConstructorOptions, NativeImage, Rectangle, WebContents } from 'electron'
import { homeChannelAccess } from './home-channel-access'
import { createLiveBridgeToggle } from './live-bridge-toggle'
import { stringPathsCapped } from '../shared/home-paths'
import {
  bridgeEnvDisabled,
  effectiveLiveBridgeEnabled,
  liveBridgeToggleAllowed,
} from './live-bridge-state'
import menuDocxIcon1x from './assets/menu-docx.png?asset'
import menuDocxIcon2x from './assets/menu-docx@2x.png?asset'
import menuXlsxIcon1x from './assets/menu-xlsx.png?asset'
import menuXlsxIcon2x from './assets/menu-xlsx@2x.png?asset'
import menuPptxIcon1x from './assets/menu-pptx.png?asset'
import menuPptxIcon2x from './assets/menu-pptx@2x.png?asset'
import menuPdfIcon1x from './assets/menu-pdf.png?asset'
import menuPdfIcon2x from './assets/menu-pdf@2x.png?asset'
import menuMdIcon1x from './assets/menu-md.png?asset'
import menuMdIcon2x from './assets/menu-md@2x.png?asset'
import menuHtmlIcon1x from './assets/menu-html.png?asset'
import menuHtmlIcon2x from './assets/menu-html@2x.png?asset'
import menuHomeIcon1x from './assets/menu-home.png?asset'
import menuHomeIcon2x from './assets/menu-home@2x.png?asset'
import { createI18n, isLang, normalizeLang, setUiLang, type Lang } from '@airy-office/i18n'
import {
  ALL_OPEN_EXTENSIONS,
  AUTHOR_NAME_KEY,
  DEFAULT_SAVE_DIR_KEY,
  DROP_OPEN_CHANNEL,
  COPILOT_GUIDE_URL,
  DOCS_README_URL,
  GITHUB_REPO_URL,
  atomicWriteFile,
  openHelpUrl,
  OPEN_EXTENSION_GROUPS,
  readAuthorNameSetting,
  sanitizeAuthorName,
  appMenuLabels,
  contextMenuLabels,
  editMenuTemplate,
  forgetRendererFileAccess,
  grantRendererDir,
  installContextMenu,
  installNavigationGuard,
  isRecoverableRendererCrash,
  toggleDevToolsItem,
  isUsableSaveDir,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  voidLoad,
  windowMenuTemplate,
} from '@airy-office/electron-utils'
import { readAppSettings, writeAppSetting, writeAppSettings } from './app-settings'
import {
  LAST_RUN_VERSION_KEY,
  STAR_PROMPT_KEY,
  asStarPromptState,
  isUpgradeLaunch,
  shouldShowStarPrompt,
  shouldShowUpgradeStarPrompt,
  withDocOpen,
  withFirstRun,
  withResolved,
  withShown,
} from './star-prompt'
import { handleDroppedFiles } from './dropped-files'
import { mainStrings } from './i18n/strings-main'
import { ProjectStore } from '@airy-office/project-store'

import {
  buildDocsMenu,
  configureDocsRuntime,
  docsFileRenamed,
  docsQueryDirty,
  requestDocsClose,
  setDocsOpenPathRouter,
  readRecentFiles,
  readStarredFiles,
  recordRecentFile,
  removeRecentFiles,
  removeStarredFiles,
  replaceRecentFile,
  registerAiIpc,
  registerProjectIpc,
  toggleStarredFile,
  registerDocsIpc,
  setDocsExtraFileMenuItems,
  setDocsFileMenuHeadItems,
  setDocsMenuGate,
  setDocsShellHooks,
  createAiDocument,
  projectFileRenamed,
  setDocsShellWindow,
  setDocsFileSavedHook,
  setDocsFileOpenedHook,
  setSessionPathResolver,
  defaultSaveDir,
  uniquePathIn,
} from '../../../docs/src/main/docs-main'
import { blankXlsxBuffer } from '../../../sheets/src/gateway/csv-import'
import { blankPdfBuffer } from '../../../pdf/src/main/blank-pdf'
import {
  configureSheetsRuntime,
  hasActiveQueuedWorkbook,
  installSheetsMenu,
  markSheetsShuttingDown,
  requestSheetsClose,
  resolveSheetsSessionPath,
  sendSheetsMenuAction,
  setSheetsMenuReadyHook,
  sheetsFileRenamed,
  setSheetsCloseTabHook,
  setSheetsExtraFileMenuItems,
  setSheetsFileMenuHeadItems,
  setSheetsOpenPathRouter,
  setSheetsShellWindow,
  setSheetsWorkbookOpenedHook,
  startSheetsCaptureServer,
  stopSheetsSidecar,
} from '../../../sheets/src/main/sheets-main'
import {
  configureSlidesRuntime,
  installSlidesMenu,
  replaceSlidesRecentFile,
  requestSlidesClose,
  setSlidesCloseTabHook,
  setSlidesExtraFileMenuItems,
  setSlidesFileMenuHeadItems,
  setSlidesOpenedHook,
  setSlidesOpenPathRouter,
  setSlidesShellWindow,
  setSlidesShowBleed,
  slidesFileRenamed,
} from '../../../slides/src/main/slides-main'
import {
  clearPdfDirty,
  configurePdfRuntime,
  flushPdfSave,
  pdfIsDirty,
  requestPdfClose,
  requestPdfSaveAs,
  sendPdfPrintRequest,
  setPdfRenamedHook,
  setPdfSaveAsInFlight,
} from '../../../pdf/src/main/pdf-main'
import { PDF_CHANNELS } from '../../../pdf/src/shared/ipc'
import {
  convertPdfFileToDocxLocalWithPrompt,
  disposePdfConversionWorkers,
  PdfLoadError,
} from './pdf2docx-local'
import { convertPdfFileToPptxLocalWithPrompt } from './pdf2pptx-local'
import { convertPdfFileToXlsxLocalWithPrompt } from './pdf2xlsx-local'
import { closePdfPasswordDialog, promptPdfPassword } from './pdf-password-dialog'
import {
  configureMarkdownRuntime,
  markdownFileRenamed,
  requestMarkdownClose,
  requestMarkdownSave,
  sendMarkdownExportRequest,
  sendMarkdownPrintRequest,
  setMarkdownDocxExportedHook,
  setMarkdownFileSavedHook,
} from '../../../markdown/src/main/markdown-main'
import {
  configureHtmlRuntime,
  htmlFileRenamed,
  registerHtmlSchemes,
  requestHtmlClose,
  requestHtmlSave,
  sendHtmlExportRequest,
  sendHtmlPrintRequest,
  setHtmlDocxExportPrepareHook,
  setHtmlDocxExportedHook,
  setHtmlFileSavedHook,
  setHtmlPresentHooks,
  setHtmlProvisionalTitleHook,
} from '../../../html/src/main/html-main'
import type {
  AutoSaveDefault,
  LiveBridgeEnabled,
  RecentEntry,
  RecentPage,
  RenameResult,
  StarPromptShow,
  UiTheme,
} from '../shared/home-api'
import { HOME_CHANNELS } from '../shared/home-api'
import {
  normalizeAiPanelPrefs,
  sameAiPanelPrefs,
  type AiPanelPrefs,
} from '@airy-office/ui/ai-panel-prefs'
import type { TabKind } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'
import { showErrorDialog } from './error-dialog'
import { classifyOpenFailure, reportOpenFailure, type OpenFailureDeps } from './open-failure'
import { openFromHome } from './home-open'
import {
  isInsideDirectory,
  listStagedFiles,
  orphanedStagedFiles,
  removeStagedFile,
  untitledStagingDir,
} from './untitled-staging'
import {
  isMoveTabToNewWindowInputForKind,
  switchableDigitsForKind,
  tabIndexForDigit,
  tabSwitchTargetForInput,
} from './tab-accelerators'
import {
  pruneSession,
  readSessionState,
  serializeSession,
  writeSessionState,
} from './session-state'
import {
  isWindowOnScreen,
  readWindowState,
  writeWindowState,
  type WindowState,
} from './window-state'
import { normalizeRecentQuery, pageRecentPaths, statPathEntries } from './recent-files'
import { createQueuedWorkbookDelivery } from './queued-workbook-delivery'
import { isSameFile, isValidRenameName } from './rename-validation'
import { renameFileInAllWindows } from './rename-broadcast'
import { TabManager } from './tab-manager'
import type { DetachedTab } from './tab-manager'
import { createQuitFlow } from './quit-flow'
import { createShutdownEffects } from './shutdown-effects'
import { createWindowCloseGuard } from './window-close-guard'
import { ShellWindowRegistry, type ShellWindowEntry } from './shell-windows'
import { startShellBridge, stopShellBridge } from './bridge/shell-bridge'
import { cancelPendingInstall, flushPendingInstall, initUpdater, updaterMenuItems } from './updater'
import { initShellDiagnostics } from './crash-diagnostics'

/**
 * Airy unified shell: ONE Electron app hosting the docs/sheets modules as
 * WebContentsView tabs behind a WPS-style tab strip. The first ("primary")
 * window owns the persisted geometry; "Move tab to new window" opens further
 * shell windows, each with its own strip and Home tab. The shell owns the
 * lifecycle — single-instance lock, file-association routing by extension,
 * and per-active-tab menu switching (menus follow the focused window).
 * Renderers load from each module's build output (apps/docs/out,
 * apps/sheets/out), so build those before running the shell.
 */

// ANY unpacked run (`npm run shell`, `npm run dev`, `npx electron .`) must not
// share the installed app's userData or single-instance lock — otherwise a dev
// run silently quits and forwards its argv to the running installed Airy.
// AIRY_USER_DATA: test drivers point this at a scratch dir so an
// automated instance can run alongside the dev instance (separate lock).
if (!app.isPackaged)
  app.setPath('userData', process.env.AIRY_USER_DATA ?? join(app.getPath('appData'), 'Airy Dev'))

// The product renames ("AI Office" → GenOffice → Airy) changed the userData
// path (appData/Airy, from productName); migrate old user data once — the
// most recent legacy layout wins, and only into a still-empty new dir.
if (app.isPackaged) {
  const newDir = app.getPath('userData')
  const newEmpty = !existsSync(newDir) || readdirSync(newDir).length === 0
  for (const legacyName of ['GenOffice', 'AI Office']) {
    const oldDir = join(app.getPath('appData'), legacyName)
    if (newEmpty && existsSync(oldDir)) {
      cpSync(oldDir, newDir, { recursive: true })
      break
    }
  }
}

// ---- crash diagnostics (UX-1628), BEFORE the heavy startup phases ----
// Local-only Crashpad minidumps (userData/crash-dumps, uploadToServer:false)
// and a size-capped ring log of the main process (userData/logs, with the
// existing console.error diagnostics tee-ed into it), so an "the app closed
// itself" report becomes diagnosable post-factum. Init never throws and the
// diagnostics never touch the network; see crash-diagnostics.ts.
initShellDiagnostics({ userDataDir: app.getPath('userData') })

// module build outputs: packaged builds carry them as extraResources
// (resources/modules/*, resources/native/*); dev/unpacked resolves them
// relative to apps/shell in the monorepo layout.
const SIDECAR_EXE = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const APPS_ROOT = join(app.getAppPath(), '..')
const DOCS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'docs')
  : join(APPS_ROOT, 'docs', 'out')
const SHEETS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'sheets')
  : join(APPS_ROOT, 'sheets', 'out')
const SLIDES_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'slides')
  : join(APPS_ROOT, 'slides', 'out')
const PDF_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'pdf')
  : join(APPS_ROOT, 'pdf', 'out')
const MARKDOWN_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'markdown')
  : join(APPS_ROOT, 'markdown', 'out')
const HTML_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'html')
  : join(APPS_ROOT, 'html', 'out')
const SIDECAR_BIN = app.isPackaged
  ? join(process.resourcesPath, 'native', SIDECAR_EXE)
  : join(APPS_ROOT, 'sheets', 'native', 'xlsx-engine', 'target', 'release', SIDECAR_EXE)

configureDocsRuntime({
  preloadPath: join(DOCS_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.DOCS_RENDERER_URL,
  rendererFile: join(DOCS_OUT, 'renderer', 'index.html'),
})
configureSheetsRuntime({
  preloadPath: join(SHEETS_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.SHEETS_RENDERER_URL,
  rendererFile: join(SHEETS_OUT, 'renderer', 'index.html'),
  sidecarPath: SIDECAR_BIN,
  // Generated-file opens (exports, AI-created files) resolve the ASKING
  // tab's window first (BUG-1107 focus routing); without a sender — e.g. a
  // standalone fallback — the focused window gets the tab.
  openGeneratedPath: (path, senderWcId) =>
    openGeneratedDocument(path, managerForSender(senderWcId)),
  // The sheets AI's create_document (docx/pdf/md) funnels into the docs-owned
  // creation flow, like the pdf app below; the sender's webContents id rides
  // along so the result opens in the asking tab's window (BUG-1107).
  createDocument: (request, senderWcId) => createAiDocument(request, senderWcId),
})
configureSlidesRuntime({
  preloadPath: join(SLIDES_OUT, 'preload', 'index.js'),
  rendererDevUrl: process.env.SLIDES_RENDERER_URL,
  rendererFilePath: join(SLIDES_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path, senderWcId) =>
    openGeneratedDocument(path, managerForSender(senderWcId)),
})
configurePdfRuntime({
  preloadPath: join(PDF_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.PDF_RENDERER_URL,
  rendererFile: join(PDF_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path, senderWcId) =>
    openGeneratedDocument(path, managerForSender(senderWcId)),
  createDocument: (request, senderWcId) => createAiDocument(request, senderWcId),
})
configureMarkdownRuntime({
  preloadPath: join(MARKDOWN_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.MARKDOWN_RENDERER_URL,
  rendererFile: join(MARKDOWN_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path, senderWcId) =>
    openGeneratedDocument(path, managerForSender(senderWcId)),
})
configureHtmlRuntime({
  preloadPath: join(HTML_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.HTML_RENDERER_URL,
  rendererFile: join(HTML_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path, senderWcId) =>
    openGeneratedDocument(path, managerForSender(senderWcId)),
})
// privileged-scheme registration is only legal before app ready
registerHtmlSchemes()

// ---- UI language ----
// Persisted in userData/app-settings.json so the editor modules can read the
// same file when they pick up i18n later. AIRY_LANG overrides for tests.

const APP_SETTINGS_PATH = () => join(app.getPath('userData'), 'app-settings.json')
const WINDOW_STATE_PATH = () => join(app.getPath('userData'), 'window-state.json')
const SESSION_PATH = () => join(app.getPath('userData'), 'session.json')

let uiLang: Lang | null = null

function currentLang(): Lang {
  if (uiLang) return uiLang
  if (process.env.AIRY_LANG) {
    uiLang = normalizeLang(process.env.AIRY_LANG)
    setUiLang(uiLang)
    return uiLang
  }
  const saved = readAppSettings(APP_SETTINGS_PATH()).language
  if (isLang(saved)) uiLang = saved
  uiLang ??= normalizeLang(app.getLocale())
  setUiLang(uiLang)
  return uiLang
}

function persistLang(lang: Lang): void {
  uiLang = lang
  setUiLang(lang)
  writeAppSetting(APP_SETTINGS_PATH(), 'language', lang)
}

let cachedTheme: UiTheme | null = null

function currentTheme(): UiTheme {
  if (cachedTheme) return cachedTheme
  const saved = readAppSettings(APP_SETTINGS_PATH()).theme
  cachedTheme = saved === 'light' || saved === 'dark' ? saved : 'system'
  return cachedTheme
}

let cachedAutoSaveDefault: AutoSaveDefault | null = null

function currentAutoSaveDefault(): AutoSaveDefault {
  if (cachedAutoSaveDefault) return cachedAutoSaveDefault
  const saved = readAppSettings(APP_SETTINGS_PATH())
  const updatedAt = saved.autoSaveDefaultUpdatedAt
  cachedAutoSaveDefault = {
    on: saved.autoSaveDefault === true,
    updatedAt: typeof updatedAt === 'number' && updatedAt > 0 ? updatedAt : 0,
  }
  return cachedAutoSaveDefault
}

let cachedAiPanelPrefs: AiPanelPrefs | null = null

function currentAiPanelPrefs(): AiPanelPrefs {
  if (cachedAiPanelPrefs) return cachedAiPanelPrefs
  const saved = readAppSettings(APP_SETTINGS_PATH())
  cachedAiPanelPrefs = normalizeAiPanelPrefs({
    fontSize: saved.aiPanelFontSize,
    customFontSize: saved.aiPanelCustomFontSize,
    spellcheck: saved.aiPanelSpellcheck,
  })
  return cachedAiPanelPrefs
}

// ---- author display name (comments / revision marks) ----

let cachedAuthorName: string | null = null

/** configured author name; '' means unset (editors fall back to their defaults) */
function currentAuthorName(): string {
  if (cachedAuthorName === null) cachedAuthorName = readAuthorNameSetting(APP_SETTINGS_PATH())
  return cachedAuthorName
}

// ---- first-run onboarding ----

// ---- "star us on GitHub" prompt (see star-prompt.ts for the rules) ----

const readStarPrompt = () =>
  asStarPromptState(readAppSettings(APP_SETTINGS_PATH())[STAR_PROMPT_KEY])
const writeStarPrompt = (state: ReturnType<typeof readStarPrompt>) =>
  writeAppSetting(APP_SETTINGS_PATH(), STAR_PROMPT_KEY, state)

/** set at startup when this is the first launch after an upgrade; consumed by
 * the first starPromptShouldShow query of the session */
let upgradeStarPromptPending = false

/** a granted show, cached for the session: repeated queries (React StrictMode
 * double-effects, AppFrame remounts) must return the same answer instead of
 * burning another lifetime show or flipping to a snoozed "false" */
let starPromptSessionGrant: StarPromptShow | null = null

/** every successful document open counts toward the prompt's value threshold */
function recordStarPromptDocOpen(): void {
  try {
    const state = readStarPrompt()
    const next = withDocOpen(state)
    if (next !== state) writeStarPrompt(next)
  } catch {
    // settings write failures must never break opening a document
  }
}

// Stargazer count for the settings About pane; fetched main-side (the
// renderer CSP has no api.github.com) and cached per session — the exact
// number is decoration, staleness is fine.
let cachedGithubStars: number | null = null

async function fetchGithubStars(): Promise<number | null> {
  if (cachedGithubStars !== null) return cachedGithubStars
  try {
    const response = await fetch('https://api.github.com/repos/besliky/airy', {
      headers: { Accept: 'application/vnd.github+json' },
      signal: AbortSignal.timeout(5000),
    })
    if (!response.ok) return null
    const body: unknown = await response.json()
    const count = (body as { stargazers_count?: unknown }).stargazers_count
    if (typeof count !== 'number' || !Number.isFinite(count)) return null
    cachedGithubStars = count
    return count
  } catch {
    return null
  }
}

const tMain = createI18n(mainStrings)

const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(currentLang(), key, params)

// ---- the shell windows + their tab managers ----
// The shell hosts one primary window and any number of secondary ones ("Move
// tab to new window"). Each window owns its TabManager and its Home renderer;
// user actions that used to resolve against the single window (menus, dialogs,
// file-open routing) resolve against the FOCUSED shell window, falling back to
// the first one alive.

const shellWindows = new ShellWindowRegistry()

function shellEntries(): readonly ShellWindowEntry[] {
  return shellWindows.list()
}

/** focused shell window, else the first alive one; null before any exists */
function focusedShellWindow(): BrowserWindow | null {
  return shellWindows.focused()?.win ?? null
}

/** the focused window's tab manager, else the first alive one */
function focusedManager(): TabManager | null {
  return shellWindows.focused()?.manager ?? null
}

/** the manager whose strip holds this editor webContents (tabs live in exactly one window) */
function managerForWebContents(webContentsId: number): TabManager | null {
  return shellWindows.managerForWebContents(webContentsId)
}

/** Routing target for a sender-identified hook call (editor module asking the
 *  shell to open/list/close something): the sender's own window when its tab
 *  is live, else the focused one (see ShellWindowRegistry.managerForSender) */
function managerForSender(senderWcId?: number): TabManager | null {
  return shellWindows.managerForSender(senderWcId)
}

function isShellWindow(win: BrowserWindow): boolean {
  return shellWindows.forWindow(win) !== undefined
}

// ---- process-level safety net ----
// All six apps' main code shares this one process, so an unhandled rejection
// or a synchronous throw outside a handler used to kill every open document
// tab. Rejections are logged only (they are common and mostly harmless; a
// dialog per rejection would spam); uncaught exceptions additionally surface
// through the single-flight error dialog and the app tries to keep running.

process.on('unhandledRejection', (reason) => {
  console.error('[shell] unhandled rejection:', reason)
})

process.on('uncaughtException', (err) => {
  console.error('[shell] uncaught exception:', err)
  showErrorDialog(focusedShellWindow(), tm('errUnhandledException'), err)
})

app.on('child-process-gone', (_event, details) => {
  // utility/network/GPU children: log for diagnostics; Chromium restarts them itself
  console.error(`[shell] child process gone: type=${details.type} reason=${details.reason}`)
})

/** Prompt the user after a tab renderer crashed; Reload restarts it, Close drops it. */
function promptRendererCrash(
  entry: ShellWindowEntry,
  info: {
    id: string
    title: string
    reason: string
  },
): void {
  const manager = entry.manager
  if (!entry.win || entry.win.isDestroyed()) return
  const isHome = info.id === 'home'
  void dialog
    .showMessageBox(entry.win, {
      type: 'error',
      message: tm('dlgRendererCrashed'),
      detail: tm('dlgRendererCrashedDetail', { reason: info.reason }),
      buttons: [tm('btnReload'), ...(isHome ? [] : [tm('btnCloseTab')]), tm('btnCancel')],
      defaultId: 0,
      cancelId: isHome ? 1 : 2,
    })
    .then(({ response }) => {
      if (response === 0) {
        // Home is the shell window's own renderer, not a manager view
        if (isHome) reloadHomeRenderer(entry)
        else manager.reloadTab(info.id)
      } else if (response === 1 && !isHome) {
        void manager.closeTab(info.id)
      }
    })
    .catch(() => undefined)
}

/** (Re)load a shell window's own Home renderer — the app bundle or dev server,
 *  NOT webContents.reload(): after a crash the webContents shows the error
 *  page, and reloading that would just redraw the error page. */
function reloadHomeRenderer(entry: ShellWindowEntry): void {
  entry.homeRendererCrashed = false
  const win = entry.win
  if (!win || win.isDestroyed()) return
  const load = process.env.ELECTRON_RENDERER_URL
    ? win.loadURL(process.env.ELECTRON_RENDERER_URL)
    : win.loadFile(join(__dirname, '../renderer/index.html'))
  load.catch((err: unknown) => {
    console.error('[shell] renderer load failed:', err)
    showErrorDialog(win, tm('dlgLoadFailed'), err)
  })
}

/** In-window error page for the crashed Home renderer — the tabs' crash page
 *  plus a Reload button. Cancel on the crash prompt then still leaves an
 *  explanatory page with a way back (View > Reload is dev-only). The button
 *  reaches the main process through the preload's dedicated airyHomeCrash
 *  channel; body and label are main-side localized strings, HTML-escaped —
 *  nothing renderer-controlled is interpolated. */
function homeCrashErrorPageUrl(): string {
  const esc = (text: string): string =>
    text.replace(/[<>&"']/g, (c) =>
      c === '<'
        ? '&lt;'
        : c === '>'
          ? '&gt;'
          : c === '&'
            ? '&amp;'
            : c === '"'
              ? '&quot;'
              : '&#39;',
    )
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>
html,body{height:100%;margin:0}
body{display:flex;align-items:center;justify-content:center;background:#f5f5f5;color:#333;
font:14px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif;text-align:center;padding:24px}
button{margin-top:16px;padding:8px 22px;font:inherit;border:1px solid #bbb;border-radius:6px;
background:#fff;color:#333;cursor:pointer}
button:hover{background:#f0f0f0}
</style></head><body><div>${esc(tm('crashPageBody'))}
<button onclick="window.airyHomeCrash&&window.airyHomeCrash.reload()">${esc(tm('btnReload'))}</button>
</div></body></html>`
  return `data:text/html;charset=utf-8,${encodeURIComponent(html)}`
}

// The crashed-Home error page's Reload button. Fire-and-forget channel: only
// the asking window's own webContents is accepted, and only while a Home
// crash is actually pending in it (the flag doubles as the guard).
ipcMain.on(HOME_CHANNELS.crashReload, (event) => {
  const entry = shellWindows.forHomeWebContents(event.sender.id)
  if (!entry || entry.win.isDestroyed() || !entry.homeRendererCrashed) return
  reloadHomeRenderer(entry)
})

/**
 * When the user creates a file from a specific project view, remember which
 * project the next save should belong to. key: 'doc' | 'sheet' | 'slide', value: projectId.
 * Consumed by each app's saveHook once the file first hits disk (P1 item 3).
 */
const pendingNewFileProject = new Map<string, string>()

/**
 * P1: after a file first hits disk, if a pending project was set earlier via
 * "create from project view", move the new file into that project automatically.
 * Called from createShellWindow's opened/saved hooks.
 */
function applyPendingProject(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase()
  let key: string | undefined
  if (ext === 'docx') key = 'doc'
  else if (ext === 'xlsx' || ext === 'xlsm' || ext === 'xls' || ext === 'csv') key = 'sheet'
  else if (ext === 'pptx') key = 'slide'
  else if (ext === 'md' || ext === 'markdown') key = 'markdown'
  else if (ext === 'html' || ext === 'htm') key = 'html'
  else if (ext === 'pdf') key = 'pdf'
  if (!key) return
  const projectId = pendingNewFileProject.get(key)
  if (!projectId) return
  pendingNewFileProject.delete(key)
  try {
    const store = new ProjectStore(app.getPath('userData'))
    store.ensureDefaultProject()
    store.resolveProjectForFile(filePath) // assign to default first (idempotent)
    store.moveFileToProject(filePath, projectId)
  } catch (err) {
    console.warn('[shell] applyPendingProject failed:', err)
  }
}

/** kind of the tab the application menu was last built for (updater status
 * changes rebuild the same menu so the Check-for-Updates label stays live) */
let currentMenuKind: TabKind = 'home'

function applyMenuFor(kind: TabKind): void {
  currentMenuKind = kind
  switch (kind) {
    case 'docs':
      buildDocsMenu()
      break
    case 'sheets':
      installSheetsMenu()
      break
    case 'slides':
      installSlidesMenu()
      break
    case 'pdf':
      buildPdfMenu()
      break
    case 'markdown':
      buildMarkdownMenu()
      break
    case 'html':
      buildHtmlMenu()
      break
    default:
      buildHomeMenu()
  }
}

/**
 * Restore the persisted window geometry, discarding state that is malformed
 * or no longer fully on a current display (monitor unplugged / resolution
 * changed) — those cases fall back to the default centered-ish window.
 */
function restoreWindowState(): WindowState | null {
  const state = readWindowState(WINDOW_STATE_PATH())
  if (
    state &&
    isWindowOnScreen(
      state,
      screen.getAllDisplays().map((d) => d.bounds),
    )
  )
    return state
  return null
}

/** capture the current geometry (normal bounds + flags) and persist it */
function persistWindowState(win: BrowserWindow): void {
  if (win.isDestroyed()) return
  try {
    // normalBounds: the restore size while maximized/fullscreen, so a relaunch
    // reopens un-maximized at the size the user had before maximizing
    const b = win.getNormalBounds()
    writeWindowState(WINDOW_STATE_PATH(), {
      x: b.x,
      y: b.y,
      width: b.width,
      height: b.height,
      isMaximized: win.isMaximized(),
      isFullScreen: win.isFullScreen(),
    })
  } catch (err) {
    // geometry persistence must never break closing or moving the window
    console.warn('[shell] window state save failed:', err)
  }
}

// ---- session persistence (tab set + active tab, restored on launch) ----

/** "Restore previous session" preference (app-settings.json `restoreSession`); absent = on */
function sessionRestoreEnabled(): boolean {
  return readAppSettings(APP_SETTINGS_PATH()).restoreSession !== false
}

let sessionSaveTimer: ReturnType<typeof setTimeout> | null = null

/** write the live windows' session in window order (skipStaged drops
 *  untitled-staged tabs — quit only; exclude leaves a closing window out;
 *  windows whose close already went through never serialize — BUG-1218).
 *  Returns whether the snapshot landed (the failure is logged here and must
 *  never break tab operations — BUG-1224: callers with persist-once
 *  bookkeeping re-arm their retry on a false return). */
function persistSessionState(skipStaged = false, exclude?: ShellWindowEntry): boolean {
  const entries = shellWindows.persistableEntries(exclude)
  if (entries.length === 0) return true
  try {
    const stagingDir = skipStaged ? UNTITLED_STAGING_DIR() : null
    const focusedIndex = (() => {
      const focused = shellWindows.focused()
      const idx = focused ? entries.indexOf(focused) : -1
      return idx >= 0 ? idx : 0
    })()
    writeSessionState(
      SESSION_PATH(),
      serializeSession(
        entries.map((entry) => {
          const tabs = stagingDir
            ? entry.manager
                .sessionTabs()
                .map((tab) =>
                  tab.filePath && isInsideDirectory(stagingDir, tab.filePath)
                    ? { ...tab, filePath: undefined }
                    : tab,
                )
            : entry.manager.sessionTabs()
          return { tabs, activeId: entry.manager.activeTabId() }
        }),
        focusedIndex,
      ),
    )
  } catch (err) {
    // session persistence must never break tab operations
    console.warn('[shell] session state save failed:', err)
    return false
  }
  return true
}

/** Persist the open-tab set (debounced — every open/close/reorder/activation fires this).
 *  No-op while quitting (BUG-1105): the quit writes ONE snapshot at the first
 *  confirmed window close; a debounced write firing later — e.g. from a dirty
 *  guard activating a tab in another window — would re-serialize from the
 *  surviving windows only and drop the already-closed ones from the session. */
function scheduleSessionSave(): void {
  if (sessionSaveTimer) clearTimeout(sessionSaveTimer)
  sessionSaveTimer = setTimeout(() => {
    sessionSaveTimer = null
    if (quitFlow.quitting) return
    persistSessionState()
  }, 800)
}

/**
 * Reopen the file-backed tabs from the previous run (quit or crash), window by
 * window in the saved order: the first saved window reuses the primary shell
 * window, every later one with surviving tabs gets its own window, cascaded
 * from the previously restored one (BUG-1223: they used to all anchor at the
 * primary's bounds and open exactly on top of each other). Files that
 * no longer exist are skipped silently; restored tabs go through the same
 * routing as a manual open (recents, dedupe, renderer read grants). Returns
 * how many tabs were restored.
 */
function restorePreviousSession(): number {
  if (!sessionRestoreEnabled()) return 0
  const saved = readSessionState(SESSION_PATH())
  if (!saved || saved.windows.length === 0) return 0
  const live = pruneSession(saved, (path) => existsSync(path))
  let opened = 0
  // the cascade source walks with the windows actually created (BUG-1223):
  // anchoring every secondary window at the primary's bounds opened them
  // exactly on top of each other; each new window cascades from the previous
  // one, like the "Move to New Window" path
  let cascadeFrom = primaryBounds()
  live.windows.forEach((window, index) => {
    if (window.tabs.length === 0) return
    const entry = index === 0 ? shellEntries()[0] : createShellWindow({ cascadeFrom })
    if (!entry) return
    if (!entry.win.isDestroyed()) {
      const bounds = entry.win.getNormalBounds()
      if (bounds.width > 0 && bounds.height > 0) cascadeFrom = bounds
    }
    for (const tab of window.tabs) {
      if (routeDocumentPath(tab.path, entry.manager)) opened++
    }
    // re-activate the tab that was active at close (the last open already left
    // its own tab active when the saved active entry could not be restored)
    if (window.activePath) {
      const activeTab = window.tabs.find((tab) => tab.path === window.activePath)
      if (activeTab) {
        const id = entry.manager.findTabIdByPath(activeTab.kind, activeTab.path)
        if (id) entry.manager.activateTab(id)
      }
    }
  })
  // bring back the window the user had focused last
  const focusedEntry = shellEntries()[Math.min(live.focusedWindow, shellEntries().length - 1)]
  if (focusedEntry && shellEntries().length > 1) focusedEntry.win.focus()
  return opened
}

/** the primary window's normal bounds — cascade source for restored windows */
function primaryBounds(): Rectangle {
  const primary = shellEntries()[0]
  if (primary && !primary.win.isDestroyed()) {
    const b = primary.win.getNormalBounds()
    if (b.width > 0 && b.height > 0) return b
  }
  return { x: 80, y: 80, width: 1200, height: 800 }
}

/** options for creating a shell window */
interface CreateShellWindowOptions {
  /** the first window: restores and owns the persisted geometry */
  primary?: boolean
  /** a tab lifted out of another window, adopted into this one's strip */
  adopt?: DetachedTab
  /** bounds to cascade a secondary window from (defaults cascade from the primary) */
  cascadeFrom?: Rectangle
}

/** cascade a secondary window from a source rectangle, clamped into its display */
function cascadedBounds(from: Rectangle): Rectangle {
  const area = screen.getDisplayMatching(from).workArea
  const width = Math.min(Math.max(from.width, 720), area.width)
  const height = Math.min(Math.max(from.height, 550), area.height)
  const x = Math.min(Math.max(from.x + 32, area.x), area.x + area.width - width)
  const y = Math.min(Math.max(from.y + 32, area.y), area.y + area.height - height)
  return { x, y, width, height }
}

function createShellWindow(options: CreateShellWindowOptions = {}): ShellWindowEntry | null {
  const primary = options.primary === true
  const saved = primary ? restoreWindowState() : null
  const cascade = !primary && !saved ? cascadedBounds(options.cascadeFrom ?? primaryBounds()) : null
  const win = new BrowserWindow({
    ...(saved
      ? { x: saved.x, y: saved.y, width: saved.width, height: saved.height }
      : cascade
        ? { x: cascade.x, y: cascade.y, width: cascade.width, height: cascade.height }
        : { width: 1360, height: 900 }),
    minWidth: 720,
    minHeight: 550,
    title: 'Airy',
    // vibrancy: editor modules punch translucent regions (e.g. the slides
    // thumbnail pane) through to the desktop
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const, vibrancy: 'sidebar' as const }
      : {}),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  // flags apply after creation (constructor options cannot restore them);
  // maximized bounds are laid out by the OS, the tab strip follows via resize
  if (saved?.isFullScreen) win.setFullScreen(true)
  else if (saved?.isMaximized) win.maximize()
  const shellWcId = win.webContents.id

  // The entry is referenced by closures created before it exists (the crash
  // UI hands the manager its window) — resolve it through a holder instead.
  const entryRef: { entry?: ShellWindowEntry } = {}
  const manager = new TabManager(
    win,
    () => {
      win.webContents.send(TABS_CHANNELS.changed, manager.list())
      scheduleSessionSave()
    },
    applyMenuFor,
    // no extension: these tabs have no file on disk yet; the title becomes the
    // real filename (the localized untitled default + .docx etc.) once the first save lands
    (kind) =>
      kind === 'docs'
        ? tm('untitledDoc')
        : kind === 'slides'
          ? tm('untitledDeck')
          : kind === 'markdown'
            ? tm('untitledMarkdown')
            : kind === 'html'
              ? tm('untitledHtml')
              : tm('untitledSheet'),
    // renderer-crash recovery: in-tab error page + localized Reload/Close prompt
    {
      errorPageBody: () => tm('crashPageBody'),
      onCrash: (info) => {
        const entry = entryRef.entry
        if (entry) promptRendererCrash(entry, info)
      },
    },
  )
  const entry: ShellWindowEntry = {
    win,
    manager,
    homeWebContentsId: shellWcId,
    homeRendererCrashed: false,
    primary,
  }
  entryRef.entry = entry
  shellWindows.add(entry)
  // "move tab to a new window" (accelerator inside an editor view of THIS window)
  manager.onMoveTabToNewWindow = (id) => moveTabToNewWindow(entry, id)
  // a tab closed without ever saving its staged untitled file — delete the
  // scratch file (quit goes through the window close path instead, not here)
  manager.onTabClosed = (tab) => {
    if (tab.filePath) removeStagedTabFile(tab.filePath)
  }

  // Only the primary window owns the persisted geometry: a second window
  // cascades and does not clobber the saved bounds.
  if (primary) {
    let geometrySaveTimer: ReturnType<typeof setTimeout> | null = null
    const scheduleGeometrySave = (): void => {
      if (geometrySaveTimer) clearTimeout(geometrySaveTimer)
      geometrySaveTimer = setTimeout(() => {
        geometrySaveTimer = null
        persistWindowState(win)
      }, 500)
    }
    win.on('resize', scheduleGeometrySave)
    win.on('move', scheduleGeometrySave)
    win.on('maximize', () => persistWindowState(win))
    win.on('unmaximize', () => persistWindowState(win))
    win.on('enter-full-screen', () => persistWindowState(win))
    win.on('leave-full-screen', () => persistWindowState(win))
    win.on('close', () => {
      if (geometrySaveTimer) clearTimeout(geometrySaveTimer)
      persistWindowState(win)
    })
  }
  // dragging the window by the tab strip's blank (draggable) area produces no
  // DOM event anywhere — will-move is the only signal to dismiss popovers
  win.on('will-move', () => broadcastChromePressed())
  // This window claims the process-global menu/active-editor targets while
  // focused (a detached editor window or another shell window may hold them);
  // focus coming back re-points everything at THIS window's active tab.
  win.on('focus', () => {
    shellWindows.notifyFocused(win)
    manager.refreshActiveTargets()
    bindWindowGlobals(entry)
  })
  // Ctrl/Cmd+1..8 → tab N, Ctrl/Cmd+9 → last tab, Ctrl/Cmd+Shift+K → move the
  // active tab to a new window — for keydowns in this window's own (Home)
  // renderer. The shell-built menus cannot carry the digits when an editor
  // owns the menu bar. Editor tabs are sibling WebContentsView whose keydowns
  // never reach this hook; TabManager attaches the same decisions to every
  // editor view (see watchTabAccelerators).
  win.webContents.on('before-input-event', (event, input) => {
    // docs tabs reserve the chord for Word's small-caps mnemonic (see
    // MOVE_TAB_TO_NEW_WINDOW_RESERVED_KINDS), like reserved digits
    const active = manager.list().find((t) => t.active)
    if (isMoveTabToNewWindowInputForKind(input, active?.kind)) {
      event.preventDefault()
      if (active && active.id !== 'home') moveTabToNewWindow(entry, active.id)
      return
    }
    const target = tabSwitchTargetForInput(input, manager.list())
    if (target === null) return
    event.preventDefault()
    manager.activateTab(target)
  })

  // The Home tab is this window's own renderer: same crash recovery as editor
  // tabs — the in-tab error page loads BEFORE the prompt, so Cancel still
  // leaves an explanatory page (with its own Reload button) instead of a dead
  // chrome shell; Reload restarts the renderer.
  win.webContents.on('render-process-gone', (_event, details) => {
    if (!isRecoverableRendererCrash(details.reason) || entry.homeRendererCrashed) return
    entry.homeRendererCrashed = true
    voidLoad(win.webContents.loadURL(homeCrashErrorPageUrl()), 'home crash error page')
    promptRendererCrash(entry, { id: 'home', title: 'Airy', reason: details.reason })
  })

  // Closing a window walks its dirty sheets/pdf/slides/docs tabs through the
  // same save/don't-save/cancel prompt; any cancel aborts the close (and an
  // in-flight quit with it). A close event arriving while a prompt is still
  // open is swallowed — the running cycle owns the decision (BUG-1217).
  // docs dirtiness lives renderer-side, so any live docs tab forces the async
  // path and gets queried there (clean tabs pass through without activation).
  win.on(
    'close',
    createWindowCloseGuard({
      // walk order: sheets, pdf, markdown, html, slides, then docs
      dirtyTabs: () => [
        ...tagGuardTabs('sheets', manager.dirtySheetsTabs()),
        ...tagGuardTabs('pdf', manager.dirtyPdfTabs()),
        ...tagGuardTabs('markdown', manager.dirtyMarkdownTabs()),
        ...tagGuardTabs('html', manager.dirtyHtmlTabs()),
        ...tagGuardTabs('slides', manager.dirtySlidesTabs()),
        ...tagGuardTabs('docs', manager.docsTabs()),
      ],
      requestClose: async (tab) => {
        // any false below aborts this window's close — and with it the whole
        // quit when one was in flight (abortAppQuit unwinds the quit state)
        switch (tab.kind) {
          case 'sheets':
            manager.activateTab(tab.id)
            return requestSheetsClose(tab.webContents, win)
          case 'pdf':
            manager.activateTab(tab.id)
            return requestPdfClose(tab.webContents, win)
          case 'markdown':
            manager.activateTab(tab.id)
            return requestMarkdownClose(tab.webContents, win)
          case 'html':
            manager.activateTab(tab.id)
            return requestHtmlClose(tab.webContents, win)
          case 'slides':
            manager.activateTab(tab.id)
            return requestSlidesClose(tab.webContents, win)
          case 'docs':
            if (!(await docsQueryDirty(tab.webContents))) return true
            manager.activateTab(tab.id)
            return requestDocsClose(tab.webContents, win)
        }
      },
      finishClose: () => finishWindowClose(entry),
      closeWindow: () => win.close(),
      isWindowAlive: () => !win.isDestroyed(),
      abortQuit: abortAppQuit,
      logFailure: (err) => console.error('[shell] window close guard failed:', err),
    }),
  )

  win.on('closed', () => {
    shellWindows.remove(win)
    forgetRendererFileAccess(shellWcId)
  })

  // A rejected load used to become an unhandled rejection; surface it instead
  // (single-flight error dialog) so a missing/corrupt bundle is visible.
  reloadHomeRenderer(entry)
  bindWindowGlobals(entry)
  if (options.adopt) manager.adoptTab(options.adopt)
  return entry
}

/** set by before-quit: close events during a quit must not persist the
 *  session window-by-window (the last window's write would drop the earlier
 *  ones); one write at the first confirmed close keeps them all. A cancelled
 *  dirty-guard unwinds the quit (abortAppQuit) so later ordinary closes keep
 *  their per-window semantics (BUG-1104). */
const quitFlow = createQuitFlow()

/** before-quit's process-wide effects (sheets no-prompt close mode, sidecar,
 *  live bridge, pdf workers) and their rollback for a cancelled quit
 *  (BUG-1216: they used to stay armed forever after one aborted quit) */
const shutdownEffects = createShutdownEffects({
  setSheetsShuttingDown: markSheetsShuttingDown,
  stopSidecar: stopSheetsSidecar,
  disposePdfWorkers: disposePdfConversionWorkers,
  bridgeEnabled: liveBridgeEnabled,
  stopBridge: stopShellBridge,
  // BUG-1312: the rollback restart rides the same serialized chain as the
  // Settings toggle (a plain startLiveBridge() could interleave with a
  // concurrent toggle-off: stop settles first, the bridge ends up up while
  // liveBridge=false). The setting is re-read inside the chain so a toggle
  // that landed while this restart waited keeps the bridge down.
  startBridge: async () => {
    await setLiveBridgeEnabled.runExclusive(async () => {
      if (liveBridgeEnabled()) await startLiveBridge()
    })
  },
  log: (message, err) => console.error(message, err),
})

/**
 * Bookkeeping when a window's close is really going through (both the clean
 * path and the post-guard path): its never-saved untitled tabs die with it —
 * delete their staged scratch files. The session is then written from the
 * surviving windows (staged tabs included: they are still live and a later
 * crash must keep restoring them). During an app-wide quit one write keeps
 * every window's file-backed tabs for the next launch (staged ones dropped);
 * closing the last window without quitting keeps the old single-window quit
 * semantics — this window's file-backed tabs survive in the session.
 */
function finishWindowClose(entry: ShellWindowEntry): void {
  const stagingDir = UNTITLED_STAGING_DIR()
  for (const tab of entry.manager.sessionTabs()) {
    if (tab.filePath && isInsideDirectory(stagingDir, tab.filePath))
      removeStagedTabFile(tab.filePath)
  }
  const decision = quitFlow.closeDecision(shellEntries().length)
  if (decision.persist) {
    // a failed quit-time snapshot write must not count as the one write
    // (BUG-1224): quitSessionPersisted is already true, so without the
    // re-arm every later confirmed close would skip its retry and the next
    // launch would resurrect windows closed before the failure (disk full,
    // userData gone read-only, …) from the stale snapshot
    let written = decision.excludeClosing
      ? persistSessionState(false, entry)
      : persistSessionState(decision.skipStaged)
    if (!written) {
      // BUG-1307: on the LAST confirmed close there is no successor close
      // the re-arm could retry into, so retry in place — the closing window
      // is still registered here (closingConfirmed is set below), making
      // this the only point where the exact intended snapshot is writable
      written = decision.excludeClosing
        ? persistSessionState(false, entry)
        : persistSessionState(decision.skipStaged)
    }
    if (!written) quitFlow.markSnapshotWriteFailed()
  }
  // mark only AFTER this window's own snapshot write (the quit snapshot and
  // the last-window ordinary close legitimately include the closer): every
  // write from here on — another window's ordinary close, the recovery
  // rewrite when a later guard cancels the quit — must leave this window
  // out; 'closed' removes it from the registry moments later (BUG-1218: the
  // recovery rewrite used to resurrect a window whose confirmed close had
  // not yet received its asynchronous 'closed' event)
  entry.closingConfirmed = true
}

/**
 * A cancelled dirty-guard stopped this window's close; when that close was
 * part of an app-wide quit, the quit is aborted with it (the other windows
 * keep running). Unwind the quit bookkeeping — and if a quit-time session
 * write already landed, re-serialize from the live windows so the aborted
 * quit leaves session state as if it never happened (BUG-1104: the armed
 * flag used to turn every later close into a skipped write). The quit's
 * before-quit effects unwind with it (BUG-1216): sheets close prompts come
 * back and the live bridge restarts when its setting says so.
 */
function abortAppQuit(): void {
  if (!quitFlow.quitting) return
  const restoreSession = quitFlow.cancel()
  shutdownEffects.rollback()
  // a cancelled guard also disarms an updater install that was riding this
  // quit (BUG-411): the staged update waits for a later natural quit instead
  // of installing under the still-running app
  cancelPendingInstall()
  if (restoreSession) {
    // BUG-1311: a failed rewrite leaves the quit-time snapshot on disk, and
    // the next launch resurrects windows whose close had confirmed before the
    // cancel (the BUG-1218 class). Retry once — a transient failure (AV scan,
    // momentary lock) recovers — and escalate loudly when the retry fails too,
    // so a stale snapshot is at least diagnosable instead of silently wrong.
    if (!persistSessionState(false) && !persistSessionState(false)) {
      console.error(
        '[shell] recovery rewrite after the aborted quit failed twice; ' +
          'the stale quit-time snapshot stays on disk and may restore ' +
          'windows the user had already closed',
      )
    }
  }
}

/** which editor's close prompt runs for a dirty-tab walk step */
type GuardedTabKind = 'sheets' | 'pdf' | 'markdown' | 'html' | 'slides' | 'docs'

/** a TabManager dirty-tab collector result tagged with its prompt kind */
interface KindTaggedGuardTab {
  id: string
  webContents: WebContents
  kind: GuardedTabKind
}

/** tag one collector's tabs for the window close guard's kind dispatch */
function tagGuardTabs(
  kind: GuardedTabKind,
  tabs: Array<{ id: string; webContents: WebContents }>,
): KindTaggedGuardTab[] {
  return tabs.map((tab) => ({ ...tab, kind }))
}

/** point the editor modules' dialog parents at this window (focus follows the shell) */
function bindWindowGlobals(entry: ShellWindowEntry): void {
  setDocsShellWindow(entry.win)
  setSheetsShellWindow(entry.win)
  setSlidesShellWindow(entry.win)
}

/**
 * "Move to New Window": lift the tab (its live editor view) out of the source
 * window's strip and open a second shell window that adopts it. The source
 * window keeps running (its Home tab always remains); the new window gets its
 * own tab strip with Home plus the moved tab, cascaded from the source.
 */
function moveTabToNewWindow(source: ShellWindowEntry, tabId: string): void {
  if (source.win.isDestroyed()) return
  const detached = source.manager.detachTab(tabId)
  if (!detached) return
  createShellWindow({
    cascadeFrom: source.win.getNormalBounds(),
    adopt: detached,
  })
}

/**
 * Wire the editor modules' process-global hooks ONCE (they are module-level
 * singletons in each editor main — re-pointing them per window would clobber
 * the previous window). Every hook resolves its target dynamically: by the
 * reporting webContents when the editor names one (the tab lives in exactly
 * one window), else against the focused shell window.
 */
function installShellModuleHooks(): void {
  // pushRecent-triggered docs menu rebuilds must not clobber the active tab's
  // menu (of whichever shell window currently owns the menu bar)
  setDocsMenuGate(() =>
    shellEntries().some((e) => e.manager.list().some((t) => t.active && t.kind === 'docs')),
  )

  setSlidesShowBleed((wc, on) => managerForWebContents(wc.id)?.setContentBleed(wc, on))
  setHtmlPresentHooks({
    setBleed: (wc, on) => managerForWebContents(wc.id)?.setContentBleed(wc, on),
    hostWindow: () => focusedShellWindow(),
    openTab: (owner, title) => {
      ;(managerForWebContents(owner.id) ?? focusedManager())?.openHtmlPresentTab(owner, title)
      return true
    },
    closeTab: (wc) => {
      const manager = managerForWebContents(wc.id)
      const id = manager?.tabIdForWebContents(wc.id)
      if (manager && id) void manager.closeTab(id)
      return !!id
    },
  })
  setDocsShellHooks({
    // every hook resolves the SENDER's window first (BUG-1107: a background
    // docs tab in an unfocused window must see its own tabs and get its
    // results there, not in whichever window holds focus); menu-driven calls
    // arrive without a sender and fall back to the focused window
    openTab: (openPath, options, senderWcId) => {
      // win:new arrives from a docs renderer with a renderer-named path:
      // route it through the same confinement an OS-level open uses
      // (grant the folder, dedupe an already-open document). Falls through
      // to a plain tab for paths the router cannot place (e.g. missing file).
      if (openPath && openDocumentPath(openPath, managerForSender(senderWcId) ?? undefined)) return
      managerForSender(senderWcId)?.openDocsTab(openPath, options)
    },
    openAiDocTab: (content, senderWcId) =>
      managerForSender(senderWcId)?.openDocsTab(undefined, { newBlank: true, aiContent: content }),
    listTabs: (senderWcId) =>
      (managerForSender(senderWcId)?.list() ?? [])
        .filter((t) => t.kind === 'docs')
        .map((t) => ({ id: t.id, title: t.title, focused: t.active })),
    focusTab: (id, senderWcId) => managerForSender(senderWcId)?.activateTab(id),
    closeActiveTab: () => focusedManager()?.closeActiveTab(),
    openGeneratedPath: (path, senderWcId) =>
      openGeneratedDocument(path, managerForSender(senderWcId)),
  })
  setSheetsCloseTabHook(() => focusedManager()?.closeActiveTab())
  // A File > Open inside an editor tab that picked a file of another type is
  // routed by extension exactly like an open from Home
  setDocsOpenPathRouter((path) => openDocumentPath(path))
  setSheetsOpenPathRouter((path) => openDocumentPath(path))
  setSlidesOpenPathRouter((path) => openDocumentPath(path))
  // ⌘W targets the focused window: in a detached slides editor window it closes
  // that window (running its own close guard), not the shell's active tab
  setSlidesCloseTabHook(() => {
    const focused = BrowserWindow.getFocusedWindow()
    if (focused && !isShellWindow(focused)) focused.close()
    else focusedManager()?.closeActiveTab()
  })
  // When ⌘O opens a file inside a tab, sync the tab title/path (used for de-dup by path) and record it as recent.
  // The first save / save-as fires this too, so applyPendingProject also runs here.
  setSheetsWorkbookOpenedHook((wc, path) => {
    const manager = managerForWebContents(wc.id)
    if (!manager) return
    // a staged untitled workbook's first save landed elsewhere — the scratch
    // file under userData is dead (moved by auto-rename, or superseded)
    const previous = manager.tabFilePathFor(wc.id)
    if (previous && previous !== path) removeStagedTabFile(previous)
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  setSlidesOpenedHook((wc, path) => {
    const manager = managerForWebContents(wc.id)
    if (!manager) return
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // docs' save-as / silent first save lands on a new path → sync the tab title too
  setDocsFileSavedHook((wc, path) => {
    const manager = managerForWebContents(wc.id)
    if (!manager) return
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // ⌘O / open-path inside a docs tab: sync the tab title immediately, same
  // contract as the sheets/slides opened hooks (a plain save to the original
  // path never renames the tab, so the open must — r115)
  setDocsFileOpenedHook((wcId, path) => {
    const manager = managerForWebContents(wcId)
    if (!manager) return
    manager.setTabFileFor(wcId, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // markdown untitled first save / Save As lands on a new path
  setMarkdownFileSavedHook((wc, path) => {
    const manager = managerForWebContents(wc.id)
    if (!manager) return
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  setHtmlFileSavedHook((wc, path) => {
    const manager = managerForWebContents(wc.id)
    if (!manager) return
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  setHtmlProvisionalTitleHook((wc, title) =>
    managerForWebContents(wc.id)?.setTabTitleFor(wc.id, title),
  )
  // pdf content-derived auto-rename: the file moved on disk, follow it everywhere
  setPdfRenamedHook((wc, oldPath, newPath) => {
    const manager = managerForWebContents(wc.id)
    if (!manager) return
    const wasStaged = isInsideDirectory(UNTITLED_STAGING_DIR(), oldPath)
    manager.setTabFileFor(wc.id, newPath)
    if (wasStaged) {
      // the rename moved the staged file into the real save folder: adopt the
      // recents entry and any pending "create in project" for the final path
      removeRecentFiles([oldPath])
      recordRecentFile(newPath)
      applyPendingProject(newPath)
    } else {
      replaceRecentFile(oldPath, newPath)
      projectFileRenamed(oldPath, newPath)
    }
  })
  // markdown "convert & open in Docs" → route the fresh .docx to a docs tab
  // in the EXPORTING tab's window (BUG-1107)
  setMarkdownDocxExportedHook((path, senderWcId) => {
    openDocumentPath(path, managerForSender(senderWcId) ?? undefined)
  })
  // Word export to a path already open in a docs tab: close that tab before the file is
  // written (its unsaved-changes prompt applies, and a later save of the stale document
  // could otherwise overwrite the export); a cancelled close aborts the export.
  setHtmlDocxExportPrepareHook(async (path, senderWcId) => {
    const manager = managerForSender(senderWcId)
    if (!manager) return true
    const stale = manager.findDocsTabByPath(path)
    if (!stale) return true
    const active = manager.list().find((t) => t.active)?.id
    await manager.closeTab(stale)
    if (active && active !== stale) manager.activateTab(active)
    return !manager.findDocsTabByPath(path)
  })
  setHtmlDocxExportedHook((path, senderWcId) => {
    openDocumentPath(path, managerForSender(senderWcId) ?? undefined)
  })
}

// ---- routing: one dispatch function for every open path ----

const DOCX_RE = /\.docx$/i
const XLSX_RE = /\.(xlsx|xlsm|xls|csv)$/i
const PPTX_RE = /\.pptx$/i
const PDF_RE = /\.pdf$/i
const MD_RE = /\.(md|markdown)$/i
const HTML_RE = /\.html?$/i

/** document formats we recognize but don't open — surfaced as a dialog, not silently dropped */
const UNSUPPORTED_DOC_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsb|pages|key|numbers)$/i

/** union of the per-module open regexes above, built from their sources so the
 *  two cannot drift (unreadable-path guard in routeDocumentPath, missing-file
 *  launch probe in missingFileIn) */
const SUPPORTED_DOC_RE = new RegExp(
  [DOCX_RE, XLSX_RE, PPTX_RE, PDF_RE, MD_RE, HTML_RE].map((re) => re.source).join('|'),
  'i',
)

/**
 * The suite-wide open-dialog filter list: one entry per document type plus the
 * combined "all supported" filter, shared by the Home browse, every shell File
 * > Open, and (in shell mode) the editors' own File > Open. Extension groups
 * come from electron-utils; the names are localized here. Legacy .doc/.ppt
 * binaries stay selectable so they surface the explicit "not supported"
 * dialog via openDocumentPath instead of being grayed out.
 */
function openDialogFilters() {
  return [
    { name: tm('filterSupported'), extensions: [...ALL_OPEN_EXTENSIONS] },
    { name: tm('filterWord'), extensions: [...OPEN_EXTENSION_GROUPS.word] },
    { name: tm('filterExcel'), extensions: [...OPEN_EXTENSION_GROUPS.excel] },
    { name: tm('filterPpt'), extensions: [...OPEN_EXTENSION_GROUPS.ppt] },
    { name: tm('filterPdf'), extensions: [...OPEN_EXTENSION_GROUPS.pdf] },
    { name: tm('filterMarkdown'), extensions: [...OPEN_EXTENSION_GROUPS.markdown] },
    { name: tm('filterHtml'), extensions: [...OPEN_EXTENSION_GROUPS.html] },
  ]
}

function supportedFileIn(argv: string[]): string | null {
  return (
    argv.find(
      (arg) =>
        (DOCX_RE.test(arg) ||
          XLSX_RE.test(arg) ||
          PPTX_RE.test(arg) ||
          PDF_RE.test(arg) ||
          MD_RE.test(arg) ||
          HTML_RE.test(arg)) &&
        existsSync(arg),
    ) ?? null
  )
}

function unsupportedFileIn(argv: string[]): string | null {
  return argv.find((arg) => UNSUPPORTED_DOC_RE.test(arg) && existsSync(arg)) ?? null
}

/**
 * argv entry naming a supported document that cannot be stat'ed — EACCES on a
 * parent directory or a file that has vanished. supportedFileIn and
 * unsupportedFileIn drop such paths (their existsSync guard also filters argv
 * junk), so an inaccessible double-click/CLI launch used to be silent
 * (BUG-1655); switch-looking arguments (always leading "-") are ignored.
 */
function missingFileIn(argv: string[]): string | null {
  return (
    argv.find((arg) => !arg.startsWith('-') && SUPPORTED_DOC_RE.test(arg) && !existsSync(arg)) ??
    null
  )
}

function notifyUnsupportedFile(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase() || basename(filePath)
  showAppWarning(tm('errUnsupportedExt', { ext }))
}

/** shell-hosted warning box; focused when a shell window exists, standalone otherwise */
function showAppWarning(message: string): void {
  const options = { type: 'warning' as const, message }
  const win = focusedShellWindow()
  if (win) {
    win.show()
    win.focus()
    void dialog.showMessageBox(win, options)
  } else {
    void dialog.showMessageBox(options)
  }
}

/**
 * Files dropped from the OS into any renderer arrive via installDropOpenBridge
 * and route through the normal File > Open pipeline; detached editor windows
 * can host the drop target, so the shell must reveal itself after opening.
 */
function registerDroppedFilesIpc(): void {
  ipcMain.on(DROP_OPEN_CHANNEL, (_event, raw: unknown) =>
    handleDroppedFiles(raw, {
      openDocumentPath,
      revealShellWindow,
      showWarning: showAppWarning,
      unsupportedMessage: (exts) => tm('errUnsupportedExt', { ext: exts.join(', ') }),
    }),
  )
}

/** the single router: extension decides which module owns the file; false = nothing opened.
 *  Opens land in the focused shell window's strip, unless `into` pins a target
 *  (session restore opens each saved window's tabs into that window). */
function openDocumentPath(filePath: string, into?: TabManager): boolean {
  const opened = routeDocumentPath(filePath, into)
  if (opened) recordStarPromptDocOpen()
  return opened
}

/**
 * Open a just-written export. Unlike File > Open, an already-open PDF tab is
 * reloaded from disk so a re-export to the same path shows the new bytes
 * instead of the previous in-memory document (which may also hold unsaved
 * annotations). In-memory edits on that tab are discarded — Save would
 * overwrite the file we just exported. `into` pins the exporting tab's
 * window (BUG-1107); without it the focused window gets the tab.
 */
function openGeneratedDocument(filePath: string, into?: TabManager | null): boolean {
  const manager = into ?? focusedManager()
  if (manager && PDF_RE.test(filePath)) {
    const existing = manager.findPdfTabByPath(filePath)
    if (existing) {
      manager.reloadTab(existing)
      manager.activateTab(existing)
      return true
    }
  }
  return openDocumentPath(filePath, into ?? undefined)
}

/** localized dialog channel for a user-intended open that produced no tab
 *  (BUG-1655); the window is resolved at call time */
const openFailureDeps: OpenFailureDeps = {
  showErrorDialog: (message, err) => showErrorDialog(focusedShellWindow(), message, err),
  openFailedMessage: (name) => tm('errOpenFailed', { name }),
}

function routeDocumentPath(filePath: string, into?: TabManager): boolean {
  const manager = into ?? focusedManager()
  if (!manager) return false
  if (SUPPORTED_DOC_RE.test(filePath)) {
    // BUG-1655: a document tab is only worth opening for a path that exists
    // AND is readable — an unreadable file (EACCES/EPERM), a vanished one
    // (ENOENT) or a directory wearing a document extension (EISDIR) used to
    // open a dead tab or fall back to Home in silence. User-intended callers
    // surface the dialog via reportOpenFailure; session restore stays
    // deliberately silent here.
    if (classifyOpenFailure(filePath)) return false
  } else if (!existsSync(filePath)) {
    return false
  }
  // every shell-routed open is user-intended: its folder becomes readable
  // for the renderer of the tab that will load it (per-sender allowlist).
  // `existing`/openXTab return tab ids; grantTabFile resolves the tab's
  // webContents and grants exactly that renderer.
  const grantTo = (tabId: string): void => manager.grantTabFile(tabId, filePath)
  const openOrActivate = (
    find: () => string | undefined,
    open: () => string,
    afterOpen?: () => void,
  ): boolean => {
    const existing = find()
    if (existing) {
      manager.activateTab(existing)
      grantTo(existing)
    } else {
      grantTo(open())
      afterOpen?.()
    }
    return true
  }
  if (DOCX_RE.test(filePath)) {
    recordRecentFile(filePath)
    return openOrActivate(
      () => manager.findDocsTabByPath(filePath),
      () => manager.openDocsTab(filePath),
    )
  }
  if (XLSX_RE.test(filePath)) {
    recordRecentFile(filePath)
    return openOrActivate(
      () => manager.findSheetsTabByPath(filePath),
      () => manager.openSheetsTab(filePath),
      () => queuedWorkbookDelivery.start(),
    )
  }
  if (PPTX_RE.test(filePath)) {
    recordRecentFile(filePath)
    return openOrActivate(
      () => manager.findSlidesTabByPath(filePath),
      // For a new tab the path goes through the pending queue; the renderer
      // consumes it after mounting
      () => manager.openSlidesTab(filePath),
    )
  }
  if (PDF_RE.test(filePath)) {
    recordRecentFile(filePath)
    return openOrActivate(
      () => manager.findPdfTabByPath(filePath),
      () => manager.openPdfTab(filePath),
    )
  }
  if (MD_RE.test(filePath)) {
    recordRecentFile(filePath)
    return openOrActivate(
      () => manager.findMarkdownTabByPath(filePath),
      () => manager.openMarkdownTab(filePath),
    )
  }
  if (HTML_RE.test(filePath)) {
    recordRecentFile(filePath)
    return openOrActivate(
      () => manager.findHtmlTabByPath(filePath),
      () => manager.openHtmlTab(filePath),
    )
  }
  notifyUnsupportedFile(filePath)
  return false
}

// ---- untitled staging (blank sheets/pdf files live under userData until their first save) ----

const UNTITLED_STAGING_DIR = () => untitledStagingDir(app.getPath('userData'))

/** delete a staged file and its recent-list entry (the file never had a real home) */
function removeStagedTabFile(path: string): void {
  if (removeStagedFile(UNTITLED_STAGING_DIR(), path)) removeRecentFiles([path])
}

/** stage a blank untitled file in userData instead of the default save folder */
function stageUntitledFile(fileName: string, bytes: Buffer | Uint8Array): string {
  const dir = UNTITLED_STAGING_DIR()
  mkdirSync(dir, { recursive: true })
  const filePath = uniquePathIn(dir, fileName)
  writeFileSync(filePath, bytes)
  return filePath
}

/** staged files that survived a crash but no open tab owns — purge at launch */
function purgeOrphanStagedFiles(): void {
  const open = shellEntries()
    .flatMap((entry) => entry.manager.sessionTabs())
    .map((tab) => tab.filePath)
    .filter((path): path is string => typeof path === 'string')
  for (const path of orphanedStagedFiles(listStagedFiles(UNTITLED_STAGING_DIR()), open)) {
    removeStagedTabFile(path)
  }
}

/**
 * "New spreadsheet" stages the blank .xlsx under userData (untitled-staging/)
 * and opens it as a regular file tab — the blank in-memory demo mode has no
 * save pipeline, so the file must exist before edits, but it must not pollute
 * the default save folder. The first save opens the Save dialog anchored in
 * the default save folder (sheets' suggestSaveAs); closing without saving
 * deletes the staged file. Falls back to the old blank tab if the write fails.
 */
async function newSheetTab(): Promise<void> {
  try {
    const filePath = stageUntitledFile(`${tm('untitledSheet')}.xlsx`, await blankXlsxBuffer())
    // the staging path itself marks the workbook untitled in sheets-main:
    // its first save opens the Save dialog anchored in the default save
    // folder, and an AI content-derived rename moves it there
    if (routeDocumentPath(filePath)) recordStarPromptDocOpen()
  } catch (err) {
    console.warn('[shell] blank workbook create failed, opening in-memory blank tab:', err)
    try {
      focusedManager()?.openSheetsTab(undefined, { newBlank: true })
    } catch (fallbackErr) {
      surfaceNewTabError(fallbackErr)
    }
  }
}

/**
 * A throw anywhere in the create-tab path (view creation, sidecar resolution,
 * renderer load) used to be swallowed by `void`-ed promises and ipc-invoke
 * rejections, so the click looked like a pure no-op — the exact "AI Sheets /
 * AI Slides do nothing" alpha report. Surface the failure instead.
 */
function surfaceNewTabError(err: unknown): void {
  console.error('[shell] new tab failed:', err)
  showErrorDialog(focusedShellWindow(), tm('errNewTabFailed'), err)
}

function newDocTab(): void {
  try {
    focusedManager()?.openDocsTab(undefined, { newBlank: true })
    // creating a document is as much a value moment as opening one
    recordStarPromptDocOpen()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newSlideTab(): void {
  try {
    focusedManager()?.openSlidesTab()
    recordStarPromptDocOpen()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newMarkdownTab(): void {
  try {
    focusedManager()?.openMarkdownTab()
    recordStarPromptDocOpen()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newHtmlTab(): void {
  try {
    focusedManager()?.openHtmlTab()
    recordStarPromptDocOpen()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

/**
 * "New PDF" stages the blank single-page .pdf under userData (untitled-staging/)
 * and opens it as a regular file tab — the PDF module has no in-memory blank
 * mode (openPdfTab requires a path). The first explicit Save opens the Save
 * dialog anchored in the default save folder and rebinds the tab; closing
 * without saving deletes the staged file.
 */
async function newPdfTab(): Promise<void> {
  try {
    const filePath = stageUntitledFile(`${tm('untitledPdf')}.pdf`, await blankPdfBuffer())
    // the staging path itself marks the pdf untitled in pdf-main (AI
    // content-derived auto-naming, first-save dialog in the shell's pdf menu)
    // A pending "create in project" intentionally stays pending: it applies to
    // the final path the first save or auto-rename picks, not this staged file
    // counts one doc-open — same as the blank workbook above
    if (routeDocumentPath(filePath)) recordStarPromptDocOpen()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

/**
 * The sheets renderer subscribes to menu actions only after Univer finishes
 * mounting (seconds on cold start), so a single 'open' can fire into the void.
 * The renderer now sends a one-time menu-ready signal when its subscription is
 * live, which flushes the queued workbook immediately; two bounded resends
 * cover a stale preload that never sends the signal (the renderer's
 * has-queued-workbook self-poll remains the safety net beyond that;
 * consumption of the queued path clears the queue entry main-side in
 * sheets-main, which stops everything). The state machine lives in
 * queued-workbook-delivery.ts, unit-tested there with injected timers.
 */
const queuedWorkbookDelivery = createQueuedWorkbookDelivery({
  // only the active tab's queue entry matters here (background tabs from a
  // multi-select Open pull their path themselves via the renderer's poll)
  sendOpen: () => sendSheetsMenuAction('open'),
  isStillWaiting: () => hasActiveQueuedWorkbook() && Boolean(focusedManager()?.findSheetsTab()),
})

setSheetsMenuReadyHook(() => queuedWorkbookDelivery.onReady())

// ---- home IPC ----

function statEntries(paths: string[]): Promise<RecentEntry[]> {
  return statPathEntries(paths, new Set(readStarredFiles()))
}

/** live-bridge effective state (app-settings.json `liveBridge`, absent = enabled; env override wins) */
function liveBridgeEnabled(): boolean {
  return effectiveLiveBridgeEnabled(readAppSettings(APP_SETTINGS_PATH()).liveBridge)
}

/** start (or return the running) live bridge server — startup, the Settings
 *  toggle, and the rollback after an aborted quit (BUG-1216) share this */
function startLiveBridge() {
  return startShellBridge({
    userDataDir: app.getPath('userData'),
    getTabManager: () => focusedManager(),
  })
}

/** serialized toggle: transition the server first, persist only on success
 *  (see live-bridge-toggle.ts); a failed start rejects the IPC result and
 *  leaves the previous stored value in place */
const setLiveBridgeEnabled = createLiveBridgeToggle({
  isEnabled: () => liveBridgeEnabled(),
  toggleAllowed: () => liveBridgeToggleAllowed(),
  persist: (on) => writeAppSetting(APP_SETTINGS_PATH(), 'liveBridge', on),
  start: async () => {
    await startLiveBridge()
  },
  stop: () => stopShellBridge(),
})

function registerHomeIpc(): void {
  // home:* channels are process-global (the shell bundles every editor's
  // main code), so only a Home tab — a shell window's own renderer — may
  // drive them; any other webContents is untrusted. Every shell window's
  // Home counts. Default-deny: every handler below is registered through
  // handleHome, which sender-checks all channels except the reads explicitly
  // exempted in home-channel-access.ts.
  const isHomeRenderer = (senderId: number): boolean =>
    shellWindows.forHomeWebContents(senderId) !== undefined
  // `...args: never[]` keeps concrete handler parameter types assignable
  const handleHome = (
    channel: string,
    handler: (event: Electron.IpcMainInvokeEvent, ...args: never[]) => unknown,
  ): void => {
    ipcMain.handle(channel, (event, ...args) => {
      if (homeChannelAccess(channel) !== 'open' && !isHomeRenderer(event.sender.id)) {
        throw new Error('Untrusted IPC sender.') // the standard untrusted-sender error
      }
      return handler(event, ...(args as never[]))
    })
  }

  handleHome(HOME_CHANNELS.getAppVersion, (): string => app.getVersion())

  handleHome(HOME_CHANNELS.recents, (_event, query: unknown): Promise<RecentPage> =>
    pageRecentPaths(readRecentFiles(), query, new Set(readStarredFiles())),
  )

  // Starred files sort by mtime, which requires stat-ing them all first; they are hand-picked and few, so this is fine
  handleHome(HOME_CHANNELS.starred, async (_event, query: unknown): Promise<RecentPage> => {
    const { offset, limit, ext } = normalizeRecentQuery(query)
    const all = (await statEntries(readStarredFiles())).sort((a, b) => b.mtimeMs - a.mtimeMs)
    const filtered = ext ? all.filter((entry) => entry.ext === ext) : all
    return {
      entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
      total: filtered.length,
      totalAll: all.length,
    }
  })

  handleHome(HOME_CHANNELS.statPaths, async (_event, paths: unknown): Promise<RecentEntry[]> => {
    // bounded per call: the renderer loads project catalogs in small chunks
    // (see renderer project-files loader), so this cap only clamps absurd
    // single payloads; a whole catalog larger than HOME_PATHS_CAP is shown
    // with an honest "{n}+ files" counter, never silently truncated
    return statEntries(stringPathsCapped(paths))
  })

  handleHome(HOME_CHANNELS.toggleStar, (_event, path: unknown) => {
    if (typeof path === 'string') toggleStarredFile(path)
  })

  handleHome(HOME_CHANNELS.openPath, (_event, path: unknown) => {
    // BUG-1677: a click on an unreadable file, a vanished one or a directory
    // used to end in silence — the false routing result was dropped. The
    // shared decision raises the same #154 dialog the launch paths show.
    openFromHome(path, {
      openDocument: openDocumentPath,
      openFailure: openFailureDeps,
    })
  })

  handleHome(HOME_CHANNELS.browse, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? focusedShellWindow()
    if (!win) return
    const result = await showOpenDialogWithMemory(
      dialog,
      win,
      {
        title: tm('dlgOpenTitle'),
        filters: openDialogFilters(),
        properties: ['openFile', 'multiSelections'],
      },
      undefined,
      event.sender.id,
    )
    if (!result.canceled) for (const path of result.filePaths) openDocumentPath(path)
  })

  handleHome(HOME_CHANNELS.newDoc, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('doc', opts.projectId)
    }
    newDocTab()
  })

  handleHome(HOME_CHANNELS.newSheet, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('sheet', opts.projectId)
    }
    void newSheetTab()
  })

  handleHome(HOME_CHANNELS.newSlide, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('slide', opts.projectId)
    }
    newSlideTab()
  })

  handleHome(HOME_CHANNELS.newMarkdown, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('markdown', opts.projectId)
    }
    newMarkdownTab()
  })

  handleHome(HOME_CHANNELS.newHtml, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('html', opts.projectId)
    }
    newHtmlTab()
  })

  handleHome(HOME_CHANNELS.newPdf, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('pdf', opts.projectId)
    }
    void newPdfTab()
  })

  handleHome(HOME_CHANNELS.removeRecent, (_event, paths: unknown) => {
    const list = stringPaths(paths)
    removeRecentFiles(list)
    // an unavailable entry's star must go with it, or the Starred view keeps
    // a dead dimmed row the recents list no longer shows
    removeStarredFiles(list.filter((p) => !existsSync(p)))
  })

  handleHome(HOME_CHANNELS.revealPath, (_event, path: unknown) => {
    if (typeof path === 'string' && existsSync(path)) shell.showItemInFolder(path)
  })

  handleHome(HOME_CHANNELS.renameFile, (_event, path: unknown, newName: unknown): RenameResult => {
    if (typeof path !== 'string' || typeof newName !== 'string')
      return { ok: false, error: tm('errBadArgs') }
    const name = newName.trim()
    if (!isValidRenameName(name)) return { ok: false, error: tm('errBadName') }
    if (!existsSync(path)) return { ok: false, error: tm('errMissing') }
    const target = join(dirname(path), name)
    if (target === path) return { ok: true, path }
    // A case-only rename (Report.pdf -> report.pdf) hits the source itself on
    // case-insensitive filesystems; only a genuinely different file blocks.
    if (existsSync(target) && !isSameFile(path, target)) {
      return { ok: false, error: tm('errExists') }
    }
    try {
      renameSync(path, target)
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : tm('errRenameFailed') }
    }
    replaceRecentFile(path, target)
    // project-store's fileMap/chatIdByPath re-key too, so AI chat history follows the file
    projectFileRenamed(path, target)
    // the slides module's own recent list switches to the new path as well (used by the start screen)
    if (/\.pptx$/i.test(target)) void replaceSlidesRecentFile(path, target)
    // open tabs sync their title/path in EVERY window — the file's identity
    // changed globally, and routing through the focused window alone left a
    // second window's tab on the stale save path (Ctrl+S would write to the
    // pre-rename name, BUG-1675); each editor then syncs its internal save
    // path and title bar
    renameFileInAllWindows(
      shellWindows.list().map((entry) => entry.manager),
      path,
      target,
      {
        docs: docsFileRenamed,
        sheets: sheetsFileRenamed,
        slides: slidesFileRenamed,
        markdown: markdownFileRenamed,
        html: htmlFileRenamed,
      },
    )
    return { ok: true, path: target }
  })

  handleHome(HOME_CHANNELS.duplicateFile, (_event, path: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return
    const ext = extname(path)
    const base = basename(path, ext)
    const dir = dirname(path)
    for (let i = 1; ; i++) {
      const target = join(dir, `${base} ${tm('copySuffix')}${i === 1 ? '' : ` ${i}`}${ext}`)
      if (existsSync(target)) continue
      copyFileSync(path, target)
      recordRecentFile(target)
      return
    }
  })

  handleHome(HOME_CHANNELS.deleteFiles, async (_event, paths: unknown) => {
    const list = stringPaths(paths)
    for (const p of list) {
      try {
        await shell.trashItem(p)
      } catch {
        // file already gone or trash unavailable; still drop it from the list
      }
    }
    removeRecentFiles(list)
    // the files were deliberately destroyed — stars must not survive as ghosts
    removeStarredFiles(list)
  })

  handleHome(HOME_CHANNELS.openTrash, () => {
    if (process.platform === 'darwin') {
      void shell.openPath(join(app.getPath('home'), '.Trash'))
    } else if (process.platform === 'win32') {
      spawn('explorer.exe', ['shell:RecycleBin'], { detached: true }).unref()
    } else {
      void shell.openPath(join(app.getPath('home'), '.local', 'share', 'Trash', 'files'))
    }
  })

  handleHome(HOME_CHANNELS.getLanguage, (): Lang => currentLang())

  handleHome(HOME_CHANNELS.setLanguage, (_event, lang: unknown) => {
    if (!isLang(lang) || lang === currentLang()) return
    persistLang(lang)
    // the switcher lives on the home page, so the home menu is the active one
    buildHomeMenu()
    installDockMenu()
    installBackToHomeItems()
    for (const wc of webContents.getAllWebContents()) wc.send('app:language-changed', lang)
  })

  handleHome(
    HOME_CHANNELS.onboardingSeen,
    (): boolean => readAppSettings(APP_SETTINGS_PATH()).onboardingSeen === true,
  )

  handleHome(HOME_CHANNELS.setOnboardingSeen, (): boolean => {
    try {
      writeAppSetting(APP_SETTINGS_PATH(), 'onboardingSeen', true)
      return true
    } catch {
      return false
    }
  })

  handleHome(HOME_CHANNELS.getTheme, (): UiTheme => currentTheme())
  // editor tabs ask via the app-wide channel (symmetric with app:get-language)
  ipcMain.handle('app:get-theme', (): UiTheme => currentTheme())

  handleHome(HOME_CHANNELS.setTheme, (_event, theme: unknown) => {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system') return
    if (theme === currentTheme()) return
    cachedTheme = theme
    writeAppSetting(APP_SETTINGS_PATH(), 'theme', theme)
    nativeTheme.themeSource = theme
    for (const wc of webContents.getAllWebContents()) wc.send('app:theme-changed', theme)
  })

  handleHome(HOME_CHANNELS.getAutoSaveDefault, (): AutoSaveDefault => currentAutoSaveDefault())
  ipcMain.handle('app:get-auto-save-default', (): AutoSaveDefault => currentAutoSaveDefault())

  // author display name: persisted like the other General settings and pushed
  // to open editors live (they stamp it on new comments / revision marks)
  handleHome(HOME_CHANNELS.getAuthorName, (): string => currentAuthorName())
  ipcMain.handle('app:get-author-name', (): string => currentAuthorName())

  handleHome(HOME_CHANNELS.setAuthorName, (_event, raw: unknown): string => {
    const next = sanitizeAuthorName(raw)
    if (next === currentAuthorName()) return next
    cachedAuthorName = next
    writeAppSetting(APP_SETTINGS_PATH(), AUTHOR_NAME_KEY, next)
    for (const wc of webContents.getAllWebContents()) wc.send('app:author-name-changed', next)
    return next
  })

  handleHome(HOME_CHANNELS.setAutoSaveDefault, (_event, on: unknown) => {
    if (typeof on !== 'boolean') return
    if (on === currentAutoSaveDefault().on) return
    const next: AutoSaveDefault = { on, updatedAt: Date.now() }
    cachedAutoSaveDefault = next
    writeAppSettings(APP_SETTINGS_PATH(), {
      autoSaveDefault: next.on,
      autoSaveDefaultUpdatedAt: next.updatedAt,
    })
    for (const wc of webContents.getAllWebContents()) wc.send('app:auto-save-default-changed', next)
  })

  // Copilot live bridge toggle: persisted in app-settings.json; flipping it
  // starts/stops the local socket server immediately (no restart needed)
  handleHome(HOME_CHANNELS.getLiveBridgeEnabled, (): LiveBridgeEnabled => liveBridgeEnabled())
  handleHome(HOME_CHANNELS.setLiveBridgeEnabled, (_event, on: unknown) => {
    if (typeof on !== 'boolean') return liveBridgeEnabled()
    return setLiveBridgeEnabled(on).catch((err: unknown) => {
      console.error(`bridge server failed to toggle ${on ? 'on' : 'off'}:`, err)
      throw err // the Settings switch snaps back; the stored value is unchanged
    })
  })
  // whether AIRY_DISABLE_BRIDGE=1 pins the bridge off (Settings shows a note)
  handleHome(HOME_CHANNELS.getLiveBridgeEnvDisabled, (): boolean => bridgeEnvDisabled())

  // session restore toggle (Settings → General): read on the next launch
  handleHome(HOME_CHANNELS.getRestoreSession, (): boolean => sessionRestoreEnabled())
  handleHome(HOME_CHANNELS.setRestoreSession, (_event, on: unknown) => {
    if (typeof on === 'boolean') writeAppSetting(APP_SETTINGS_PATH(), 'restoreSession', on)
  })

  handleHome(HOME_CHANNELS.getAiPanelPrefs, (): AiPanelPrefs => currentAiPanelPrefs())
  ipcMain.handle('app:get-ai-panel-prefs', (): AiPanelPrefs => currentAiPanelPrefs())

  handleHome(HOME_CHANNELS.setAiPanelPrefs, (_event, patch: unknown): AiPanelPrefs => {
    const prev = currentAiPanelPrefs()
    const raw =
      patch !== null && typeof patch === 'object' ? (patch as Record<string, unknown>) : {}
    // unknown/malformed fields fall back to the previous value, not the default
    const next = normalizeAiPanelPrefs({
      fontSize: 'fontSize' in raw ? raw.fontSize : prev.fontSize,
      customFontSize: 'customFontSize' in raw ? raw.customFontSize : prev.customFontSize,
      spellcheck: 'spellcheck' in raw ? raw.spellcheck : prev.spellcheck,
    })
    if (sameAiPanelPrefs(next, prev)) return prev
    cachedAiPanelPrefs = next
    writeAppSettings(APP_SETTINGS_PATH(), {
      aiPanelFontSize: next.fontSize,
      aiPanelCustomFontSize: next.customFontSize,
      aiPanelSpellcheck: next.spellcheck,
    })
    for (const wc of webContents.getAllWebContents()) wc.send('app:ai-panel-prefs-changed', next)
    return next
  })

  // effective folder where new/untitled files land; the editor mains resolve
  // the same setting themselves (configuredDefaultSaveDir via docs' defaultSaveDir)
  handleHome(HOME_CHANNELS.getDefaultSaveDir, (event): string => {
    // the default save folder is a standing user choice: readable for the
    // asking renderer (per-sender allowlist)
    grantRendererDir(defaultSaveDir(), event.sender.id)
    return defaultSaveDir()
  })

  handleHome(HOME_CHANNELS.pickDefaultSaveDir, async (event): Promise<string | null> => {
    const result = await showOpenDialogWithMemory(
      dialog,
      focusedShellWindow(),
      {
        title: tm('dlgPickSaveDir'),
        defaultPath: defaultSaveDir(),
        properties: ['openDirectory', 'createDirectory'],
      },
      undefined,
      event.sender.id,
    )
    const picked = result.filePaths[0]
    if (result.canceled || !picked) return null
    if (!isUsableSaveDir(picked)) {
      showErrorDialog(focusedShellWindow(), tm('errSaveDirUnusable'), picked)
      return null
    }
    writeAppSetting(APP_SETTINGS_PATH(), DEFAULT_SAVE_DIR_KEY, picked)
    return picked
  })

  handleHome(HOME_CHANNELS.openGitHubRepo, () => {
    // through the shared safeExternalUrl gate like every other external link
    void openHelpUrl(GITHUB_REPO_URL)
  })

  handleHome(HOME_CHANNELS.githubStars, () => fetchGithubStars())

  // returning true also counts as "shown": the renderer displays it
  // unconditionally, so no separate mark-shown round-trip is needed
  handleHome(HOME_CHANNELS.starPromptShouldShow, (): StarPromptShow => {
    if (starPromptSessionGrant) return starPromptSessionGrant
    const now = Date.now()
    const state = readStarPrompt()
    const docOpens = state.docOpens ?? 0
    // dev preview of the card without waiting out the value thresholds
    // (same pattern as AIRY_FAKE_UPDATE); nothing is recorded
    if (!app.isPackaged && process.env.AIRY_FORCE_STAR_PROMPT) return { show: true, docOpens }
    const grant = (): StarPromptShow => {
      writeStarPrompt(withShown(state, now))
      starPromptSessionGrant = { show: true, docOpens }
      return starPromptSessionGrant
    }
    // first launch after an upgrade: skip the value gates once for a
    // never-prompted user (they are a proven repeat user already)
    if (upgradeStarPromptPending) {
      upgradeStarPromptPending = false
      if (shouldShowUpgradeStarPrompt(state)) return grant()
    }
    if (!shouldShowStarPrompt(state, now)) return { show: false, docOpens }
    return grant()
  })

  handleHome(HOME_CHANNELS.starPromptAction, (_event, action: unknown) => {
    if (action !== 'starred' && action !== 'later') return
    // the card was reacted to — drop the session grant so a later query (new
    // shell window on macOS) re-evaluates the real rules (snooze / resolved)
    starPromptSessionGrant = null
    // 'later' needs no write: the display was already counted by the query
    if (action === 'starred') writeStarPrompt(withResolved(readStarPrompt()))
  })
}

function stringPaths(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((p): p is string => typeof p === 'string') : []
}

// electron-vite emits ?asset files under hashed names, which breaks nativeImage's
// automatic `@2x` sibling lookup — attach the retina representation by hand
function loadMenuIcon(path1x: string, path2x: string): NativeImage {
  const icon = nativeImage.createFromPath(path1x)
  icon.addRepresentation({ scaleFactor: 2, buffer: readFileSync(path2x) })
  return icon
}

// loaded once, not on every menu open
interface MenuIconSet {
  docx: NativeImage
  xlsx: NativeImage
  pptx: NativeImage
  pdf: NativeImage
  md: NativeImage
  html: NativeImage
  home: NativeImage
}
let menuIconCache: MenuIconSet | null = null
function menuIcons(): MenuIconSet {
  menuIconCache ??= {
    docx: loadMenuIcon(menuDocxIcon1x, menuDocxIcon2x),
    xlsx: loadMenuIcon(menuXlsxIcon1x, menuXlsxIcon2x),
    pptx: loadMenuIcon(menuPptxIcon1x, menuPptxIcon2x),
    pdf: loadMenuIcon(menuPdfIcon1x, menuPdfIcon2x),
    md: loadMenuIcon(menuMdIcon1x, menuMdIcon2x),
    html: loadMenuIcon(menuHtmlIcon1x, menuHtmlIcon2x),
    home: loadMenuIcon(menuHomeIcon1x, menuHomeIcon2x),
  }
  return menuIconCache
}

const TAB_MENU_ICON: Record<TabKind, keyof MenuIconSet> = {
  home: 'home',
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  pdf: 'pdf',
  markdown: 'md',
  html: 'html',
}

// tab views see neither DOM events nor a focus change when the user clicks the
// shell chrome — relay the press so open popovers in documents can dismiss.
// The pressed document must be excluded: it already dismissed (or is opening)
// its own popovers via its local pointerdown listeners, and the async IPC
// round-trip would otherwise close a popover that very press just opened
// (home row menus died this way: pointerdown → broadcast → menu unmounts
// before the click event ever reached the menu item).
function broadcastChromePressed(exclude?: WebContents): void {
  for (const wc of webContents.getAllWebContents()) {
    if (wc !== exclude) wc.send('app:chrome-pressed')
  }
}

function registerTabsIpc(): void {
  ipcMain.on(TABS_CHANNELS.chromePressed, (event) => broadcastChromePressed(event.sender))
  // tabs:* channels are process-global; each one answers the shell window
  // whose own (Home) renderer asked — the strip the user clicked belongs to
  // that window, not to whichever one happens to be focused.
  const entryForSender = (senderId: number): ShellWindowEntry | undefined =>
    shellWindows.forHomeWebContents(senderId)
  ipcMain.handle(
    TABS_CHANNELS.list,
    (event) => entryForSender(event.sender.id)?.manager.list() ?? [],
  )
  ipcMain.handle(TABS_CHANNELS.activate, (event, id: string) =>
    entryForSender(event.sender.id)?.manager.activateTab(id),
  )
  ipcMain.handle(TABS_CHANNELS.close, (event, id: string) =>
    entryForSender(event.sender.id)?.manager.closeTab(id),
  )
  ipcMain.handle(TABS_CHANNELS.reorder, (event, id: string, toIndex: number) => {
    if (typeof id === 'string' && Number.isInteger(toIndex))
      entryForSender(event.sender.id)?.manager.reorderTab(id, toIndex)
  })
  // "all tabs" overflow menu — native popup because the editors' WebContentsView
  // would cover any DOM dropdown the shell renderer draws below the tab strip
  ipcMain.handle(TABS_CHANNELS.showMenu, (event, x: unknown, y: unknown) => {
    const entry = entryForSender(event.sender.id)
    if (!entry) return
    const manager = entry.manager
    const menu = Menu.buildFromTemplate(
      manager.list().map((tab) => ({
        label: tab.title,
        type: 'checkbox' as const,
        checked: tab.active,
        icon: menuIcons()[TAB_MENU_ICON[tab.kind]],
        click: () => manager.activateTab(tab.id),
      })),
    )
    menu.popup({
      window: entry.win,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
  // "+" new-file menu — native for the same reason as the tab list above
  ipcMain.handle(TABS_CHANNELS.showNewMenu, (event, x: unknown, y: unknown) => {
    const entry = entryForSender(event.sender.id)
    if (!entry) return
    const menu = Menu.buildFromTemplate([
      // enabled:false so pre-Sonoma macOS / Windows (no 'header' support) degrade
      // to an inert label instead of a clickable no-op item
      { label: tm('menuSectionNew'), type: 'header', enabled: false },
      {
        label: tm('menuNewDoc'),
        icon: menuIcons().docx,
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        icon: menuIcons().xlsx,
        click: () => void newSheetTab(),
      },
      {
        label: tm('menuNewSlide'),
        icon: menuIcons().pptx,
        click: () => newSlideTab(),
      },
      {
        label: tm('menuNewMarkdown'),
        icon: menuIcons().md,
        click: () => newMarkdownTab(),
      },
      {
        label: tm('menuNewHtml'),
        icon: menuIcons().html,
        click: () => newHtmlTab(),
      },
      {
        label: tm('menuNewPdf'),
        icon: menuIcons().pdf,
        click: () => void newPdfTab(),
      },
      { type: 'separator' },
      { label: tm('menuOpen'), click: () => void openFileViaDialog() },
    ])
    menu.popup({
      window: entry.win,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
  // per-tab context menu (right-click on a strip tab) — native like the two above
  ipcMain.handle(TABS_CHANNELS.showTabMenu, (event, x: unknown, y: unknown, tabId: unknown) => {
    const entry = entryForSender(event.sender.id)
    if (!entry || typeof tabId !== 'string') return
    const manager = entry.manager
    const tab = manager.tabInfo(tabId)
    if (!tab) return
    const closable = tabId !== 'home'
    const otherTabs = manager.list().filter((t) => t.id !== 'home' && t.id !== tabId)
    const menu = Menu.buildFromTemplate([
      {
        label: tm('btnCloseTab'),
        enabled: closable,
        click: () => void manager.closeTab(tabId),
      },
      {
        label: tm('tabCloseOthers'),
        enabled: otherTabs.length > 0,
        click: () => void closeOtherTabs(manager, tabId),
      },
      {
        label: tm('tabCloseAll'),
        enabled: otherTabs.length > 0 || closable,
        click: () => void closeAllTabs(manager),
      },
      { type: 'separator' },
      {
        label: tm('tabDuplicate'),
        // untitled / in-memory / present tabs have no backing file (present
        // tabs carry no filePath either)
        enabled: !!tab.filePath,
        click: () => manager.duplicateTab(tabId),
      },
      { type: 'separator' },
      {
        label: tm('moveTabToNewWindow'),
        // Home belongs to this window's own renderer and cannot move
        enabled: closable,
        click: () => moveTabToNewWindow(entry, tabId),
      },
    ])
    menu.popup({
      window: entry.win,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
}

/** Close every document tab except the given one (context-menu Close Others),
 *  in the window whose strip asked */
async function closeOtherTabs(manager: TabManager, keepId: string): Promise<void> {
  for (const tab of manager.list()) {
    if (tab.id === 'home' || tab.id === keepId) continue
    // sequential: each close may run its own unsaved-changes prompt
    await manager.closeTab(tab.id)
  }
}

/** Close every document tab (context-menu Close All); Home always stays */
async function closeAllTabs(manager: TabManager): Promise<void> {
  for (const tab of manager.list()) {
    if (tab.id === 'home') continue
    await manager.closeTab(tab.id)
  }
}

// ---- home menu ----

/** switch to the tab a Ctrl/Cmd+digit selects (menu items + before-input-event) */
function activateTabForDigit(digit: number): void {
  const manager = focusedManager()
  if (!manager) return
  const tabs = manager.list()
  const index = tabIndexForDigit(digit, tabs.length)
  if (index === null) return
  manager.activateTab(tabs[index].id)
}

/**
 * Window menu for the shell-built tab menus, carrying the Ctrl/Cmd+1..9
 * tab-switch entries the active kind does not reserve for editor shortcuts
 * (docs/sheets menus are built by the editors; their free digits still work
 * through the before-input-event hook) and "Move Tab to New Window" — the
 * menu form of the tab context-menu entry and the Ctrl/Cmd+Shift+K chord,
 * which the same hook also covers inside editor tabs.
 */
function shellWindowMenu(): MenuItemConstructorOptions {
  const base = windowMenuTemplate(process.platform, appMenuLabels(currentLang()))
  const moveTabItem: MenuItemConstructorOptions[] = [
    { type: 'separator' },
    {
      label: tm('moveTabToNewWindow'),
      accelerator: 'CmdOrCtrl+Shift+K',
      enabled: currentMenuKind !== 'home',
      click: () => {
        const entry = shellWindows.focused()
        const active = entry?.manager.list().find((t) => t.active)
        if (entry && active && active.id !== 'home') moveTabToNewWindow(entry, active.id)
      },
    },
  ]
  const digits = switchableDigitsForKind(currentMenuKind)
  if (digits.length === 0)
    return {
      ...base,
      submenu: [...((base.submenu as MenuItemConstructorOptions[]) ?? []), ...moveTabItem],
    }
  return {
    ...base,
    submenu: [
      ...((base.submenu as MenuItemConstructorOptions[]) ?? []),
      { type: 'separator' },
      {
        label: tm('menuSelectTab'),
        submenu: digits.map((digit) => ({
          label: digit === 9 ? tm('menuSelectLastTab') : tm('menuSelectTabN', { n: digit }),
          accelerator: `CmdOrCtrl+${digit}`,
          click: () => activateTabForDigit(digit),
        })),
      },
      ...moveTabItem,
    ],
  }
}

async function openFileViaDialog(): Promise<void> {
  const win = focusedShellWindow() ?? BrowserWindow.getFocusedWindow()
  if (!win) return
  const result = await showOpenDialogWithMemory(dialog, win, {
    filters: openDialogFilters(),
    properties: ['openFile', 'multiSelections'],
  })
  if (!result.canceled) for (const path of result.filePaths) openDocumentPath(path)
}

/** File > New submenu (all six document types, no accelerators — Home keeps
 *  the one suite-wide Ctrl/Cmd+N) — shared by every shell-built menu */
function newFileSubMenu(): MenuItemConstructorOptions {
  return {
    label: tm('menuSectionNew'),
    submenu: [
      { label: tm('menuNewDoc'), icon: menuIcons().docx, click: () => newDocTab() },
      { label: tm('menuNewSheet'), icon: menuIcons().xlsx, click: () => void newSheetTab() },
      { label: tm('menuNewSlide'), icon: menuIcons().pptx, click: () => newSlideTab() },
      { label: tm('menuNewMarkdown'), icon: menuIcons().md, click: () => newMarkdownTab() },
      { label: tm('menuNewHtml'), icon: menuIcons().html, click: () => newHtmlTab() },
      { label: tm('menuNewPdf'), icon: menuIcons().pdf, click: () => void newPdfTab() },
    ],
  }
}

/**
 * View menu for the shell-built tab menus (Home had none): Chromium zoom +
 * fullscreen for every tab; reload and devtools stay dev-only — reloading a
 * dirty markdown/html tab would silently drop unsaved edits, matching docs'
 * dev-gated devtools entry.
 */
function shellViewMenu(): MenuItemConstructorOptions {
  const labels = appMenuLabels(currentLang())
  return {
    label: labels.view,
    submenu: [
      { role: 'resetZoom', label: labels.actualSize },
      { role: 'zoomIn', label: labels.zoomIn },
      { role: 'zoomOut', label: labels.zoomOut },
      { type: 'separator' },
      { role: 'togglefullscreen', label: labels.fullscreen },
      ...(app.isPackaged
        ? []
        : [
            { type: 'separator' as const },
            { role: 'reload' as const, label: labels.reload },
            { role: 'forceReload' as const, label: labels.forceReload },
            toggleDevToolsItem(labels),
          ]),
    ],
  }
}

/** File-menu tail for shell-built menus: on Windows/Linux an explicit Quit
 *  (Ctrl/Cmd+Q quits the whole app, every tab); macOS gets it from the app menu */
function quitMenuItem(): MenuItemConstructorOptions[] {
  return process.platform === 'darwin'
    ? []
    : [{ type: 'separator' }, { role: 'quit' as const, label: tm('menuQuit') }]
}

function buildHomeMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        { label: tm('menuSectionNew'), type: 'header', enabled: false },
        {
          label: tm('menuNewDoc'),
          accelerator: 'CmdOrCtrl+N',
          click: () => newDocTab(),
        },
        {
          label: tm('menuNewSheet'),
          click: () => void newSheetTab(),
        },
        { label: tm('menuNewSlide'), click: () => newSlideTab() },
        { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
        { label: tm('menuNewHtml'), click: () => newHtmlTab() },
        { label: tm('menuNewPdf'), click: () => void newPdfTab() },
        { type: 'separator' },
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        { role: 'close', label: tm('menuClose') },
        ...quitMenuItem(),
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    shellViewMenu(),
    shellWindowMenu(),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [
        ...updaterMenuItems(),
        { type: 'separator' },
        { label: tm('menuOnlineDocs'), click: () => void openHelpUrl(DOCS_README_URL) },
        { label: tm('menuCopilotGuide'), click: () => void openHelpUrl(COPILOT_GUIDE_URL) },
        { type: 'separator' },
        { label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() },
      ],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- pdf menu (pdf-main has no menu of its own; the shell owns pdf tabs, so it builds one) ----

function buildPdfMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        newFileSubMenu(),
        { type: 'separator' },
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => focusedManager()?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = focusedManager()?.activePdfTab()
            if (!tab) return
            // an untitled staged pdf: the first explicit Save picks where the
            // file should live (default save folder) and rebinds the tab
            if (tab.filePath && isInsideDirectory(UNTITLED_STAGING_DIR(), tab.filePath)) {
              void saveStagedPdfAs(tab.id)
              return
            }
            void flushPdfSave(tab.webContents)
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => void savePdfAs(),
        },
        { type: 'separator' },
        // local pdf2docx (P4): in-process PDFium wasm, no cloud counterpart
        {
          label: tm('menuExportDocx'),
          click: () => void exportPdfAsDocxLocal(),
        },
        // local pdf2pptx (P25): one slide per page, no cloud counterpart
        {
          label: tm('menuExportPptx'),
          click: () => void exportPdfAsPptxLocal(),
        },
        // local pdf2xlsx (P26): one worksheet per page, no cloud counterpart
        {
          label: tm('menuExportXlsx'),
          click: () => void exportPdfAsXlsxLocal(),
        },
        { type: 'separator' },
        {
          label: tm('menuPrint'),
          accelerator: 'CmdOrCtrl+P',
          click: () => {
            const tab = focusedManager()?.activePdfTab()
            if (tab) sendPdfPrintRequest(tab.webContents)
          },
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => focusedManager()?.closeActiveTab(),
        },
        ...quitMenuItem(),
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    shellViewMenu(),
    shellWindowMenu(),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [
        ...updaterMenuItems(),
        { type: 'separator' },
        { label: tm('menuOnlineDocs'), click: () => void openHelpUrl(DOCS_README_URL) },
        { label: tm('menuCopilotGuide'), click: () => void openHelpUrl(COPILOT_GUIDE_URL) },
        { type: 'separator' },
        { label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() },
      ],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- markdown menu (markdown-main has no menu of its own; the shell owns markdown tabs) ----

function buildMarkdownMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        newFileSubMenu(),
        { type: 'separator' },
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => focusedManager()?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = focusedManager()?.activeMarkdownTab()
            if (tab) void requestMarkdownSave(tab.webContents, 'save')
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => {
            const tab = focusedManager()?.activeMarkdownTab()
            if (tab) void requestMarkdownSave(tab.webContents, 'saveAs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuExportDocx'),
          click: () => {
            const tab = focusedManager()?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'docx')
          },
        },
        {
          label: tm('menuExportPdf'),
          click: () => {
            const tab = focusedManager()?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'pdf')
          },
        },
        {
          label: tm('menuOpenInDocs'),
          click: () => {
            const tab = focusedManager()?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'docs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuPrint'),
          accelerator: 'CmdOrCtrl+P',
          click: () => {
            const tab = focusedManager()?.activeMarkdownTab()
            if (tab) sendMarkdownPrintRequest(tab.webContents)
          },
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => focusedManager()?.closeActiveTab(),
        },
        ...quitMenuItem(),
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    shellViewMenu(),
    shellWindowMenu(),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [
        ...updaterMenuItems(),
        { type: 'separator' },
        { label: tm('menuOnlineDocs'), click: () => void openHelpUrl(DOCS_README_URL) },
        { label: tm('menuCopilotGuide'), click: () => void openHelpUrl(COPILOT_GUIDE_URL) },
        { type: 'separator' },
        { label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() },
      ],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- html menu (html-main has no menu of its own; the shell owns html tabs) ----

function buildHtmlMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        newFileSubMenu(),
        { type: 'separator' },
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => focusedManager()?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = focusedManager()?.activeHtmlTab()
            if (tab) void requestHtmlSave(tab.webContents, 'save')
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => {
            const tab = focusedManager()?.activeHtmlTab()
            if (tab) void requestHtmlSave(tab.webContents, 'saveAs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuExportDocx'),
          click: () => {
            const tab = focusedManager()?.activeHtmlTab()
            if (tab) sendHtmlExportRequest(tab.webContents, 'docx')
          },
        },
        {
          label: tm('menuExportPdf'),
          click: () => {
            const tab = focusedManager()?.activeHtmlTab()
            if (tab) sendHtmlExportRequest(tab.webContents, 'pdf')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuPrint'),
          accelerator: 'CmdOrCtrl+P',
          click: () => {
            const tab = focusedManager()?.activeHtmlTab()
            if (tab) sendHtmlPrintRequest(tab.webContents)
          },
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => focusedManager()?.closeActiveTab(),
        },
        ...quitMenuItem(),
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    shellViewMenu(),
    shellWindowMenu(),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [
        ...updaterMenuItems(),
        { type: 'separator' },
        { label: tm('menuOnlineDocs'), click: () => void openHelpUrl(DOCS_README_URL) },
        { label: tm('menuCopilotGuide'), click: () => void openHelpUrl(COPILOT_GUIDE_URL) },
        { type: 'separator' },
        { label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() },
      ],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

/**
 * Save As for pdf tabs: write pending edits to the picked path only, then open the copy.
 * Non-destructive: the original file is never written, and a cancelled dialog changes
 * nothing on disk (dialog first, no flush into the source).
 */
/** In-flight guard (same pattern as exportPdfAsDocxLocal): a re-trigger while the dialog
    or write is active must not start a second flow that overwrites the first one's
    waiter/target grant or clears its autosave pause early */
let savingPdfAs = false

async function savePdfAs(): Promise<void> {
  const shellWindow = focusedShellWindow()
  const tab = focusedManager()?.activePdfTab()
  if (!tab?.filePath || !shellWindow || savingPdfAs) return
  savingPdfAs = true
  // Pause renderer autosave for the whole flow: the dialog blurs the window, and a
  // blur-triggered autosave would write the pending edits into the original file
  setPdfSaveAsInFlight(tab.webContents, true)
  try {
    const picked = await showSaveDialogWithMemory(
      dialog,
      shellWindow,
      {
        defaultPath: tab.filePath,
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
      },
      undefined,
      tab.webContents.id,
    )
    if (picked.canceled || !picked.filePath || picked.filePath === tab.filePath) return
    if (pdfIsDirty(tab.webContents.id)) {
      // Renderer applies its pending edits onto the source bytes; the pdf main
      // process writes the result to the picked path only
      if (!(await requestPdfSaveAs(tab.webContents, picked.filePath))) return
    } else {
      // No pending edits → a byte-identical copy
      copyFileSync(tab.filePath, picked.filePath)
    }
    openDocumentPath(picked.filePath)
  } finally {
    savingPdfAs = false
    setPdfSaveAsInFlight(tab.webContents, false)
  }
}

/**
 * First explicit Save of an untitled staged pdf: pick the destination (Save
 * dialog anchored in the default save folder), write the pending edits to it,
 * rebind the tab to the real path and drop the staged scratch file. Unlike
 * savePdfAs (a non-destructive copy that opens a second tab), this REPLACES
 * the staged tab: its only "file" was scratch space under userData.
 */
async function saveStagedPdfAs(tabId: string): Promise<void> {
  const shellWindow = focusedShellWindow()
  const tab = focusedManager()?.activePdfTab()
  if (!tab || tab.id !== tabId || !tab.filePath) return
  if (!shellWindow || savingPdfAs) return
  const stagedPath = tab.filePath
  savingPdfAs = true
  setPdfSaveAsInFlight(tab.webContents, true)
  try {
    const picked = await showSaveDialogWithMemory(
      dialog,
      shellWindow,
      {
        defaultPath: join(defaultSaveDir(), basename(stagedPath)),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
      },
      undefined,
      tab.webContents.id,
    )
    if (picked.canceled || !picked.filePath) return
    if (pdfIsDirty(tab.webContents.id)) {
      // Renderer applies its pending edits onto the source bytes; the pdf main
      // process writes the result to the picked path only
      if (!(await requestPdfSaveAs(tab.webContents, picked.filePath))) return
    } else {
      // No pending edits → a byte-identical copy
      copyFileSync(stagedPath, picked.filePath)
    }
    // the edits are persisted at the picked path — the staged tab goes without
    // its unsaved-changes prompt and is replaced by the real file
    clearPdfDirty(tab.webContents.id)
    await focusedManager()?.closeTab(tab.id)
    removeStagedTabFile(stagedPath)
    applyPendingProject(picked.filePath)
    openDocumentPath(picked.filePath)
  } finally {
    savingPdfAs = false
    if (!tab.webContents.isDestroyed()) setPdfSaveAsInFlight(tab.webContents, false)
  }
}

/**
 * In-flight guard: covers the whole flow (dialogs included) so re-triggering
 * from the menu can never start a second conversion
 */
let exportingPdfDocx = false

/**
 * Export as Word for pdf tabs, fully local (pdf2docx P4): flush pending
 * edits, pick the destination, convert in-process via PDFium wasm, write the
 * file and open it in a Docs tab. No login, no credits.
 */
async function exportPdfAsDocxLocal(): Promise<void> {
  const shellWindow = focusedShellWindow()
  const tab = focusedManager()?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfDocxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    const picked = await showSaveDialogWithMemory(
      dialog,
      shellWindow,
      {
        defaultPath: tab.filePath.replace(/\.pdf$/i, '.docx'),
        filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
      },
      undefined,
      tab.webContents.id,
    )
    if (picked.canceled || !picked.filePath) return
    // If the destination is already open in a docs tab, close it first (its
    // normal unsaved-changes guard applies) so the converted file opens fresh
    // instead of leaving a stale tab whose next save would clobber the result.
    const staleTabId = focusedManager()?.findDocsTabByPath(picked.filePath)
    if (staleTabId) {
      await focusedManager()?.closeTab(staleTabId)
      focusedManager()?.activateTab(tab.id)
      if (focusedManager()?.findDocsTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    // encrypted PDFs prompt for the password (P23), looping on wrong entries;
    // null result = user cancelled the prompt → abort silently
    const pdfPath = tab.filePath
    const result = await convertPdfFileToDocxLocalWithPrompt(
      pdfPath,
      (retry) =>
        promptPdfPassword(shellWindow, {
          fileName: basename(pdfPath),
          retry,
          busy: false,
          lang: currentLang(),
          strings: {
            title: tm('pdfPwdTitle'),
            prompt: tm('pdfPwdPrompt'),
            retryPrompt: tm('pdfPwdRetryPrompt'),
            ok: tm('pdfPwdOk'),
            cancel: tm('btnCancel'),
            verifying: tm('pdfPwdVerifying'),
            label: tm('pdfPwdLabel'),
            placeholder: tm('pdfPwdPlaceholder'),
            show: tm('pdfPwdShow'),
            hide: tm('pdfPwdHide'),
          },
        }),
      (page, total) => {
        if (shellWindow && !shellWindow.isDestroyed() && total > 0) {
          shellWindow.setProgressBar(page / total)
        }
      },
    )
    if (result === null) return
    await atomicWriteFile(picked.filePath, result.docx)

    // degrade transparency (plan §7.6 dual-track split): whole scan → say so
    // once; individual image-fallback pages → name them;
    // OCR-recovered scans ('ocr') are SUCCESSES — announce the recovery (the
    // user should proofread machine-read text), never the image-export notice
    const ocrPages = result.pageResults.filter((r) => r.status === 'ocr').map((r) => r.page)
    const imagePages = result.pageResults
      .filter((r) => r.status !== 'ok' && r.status !== 'ocr')
      .map((r) => r.page)
    if (result.scannedDocument) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalScannedMsg'),
        detail: tm('pdfDocxLocalScannedDetail'),
      })
    } else if (imagePages.length > 0 && ocrPages.length > 0) {
      // mixed documents surface BOTH facts in one dialog: which pages shipped
      // as images and which carry machine-read text the user should proofread
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalDegradedMsg'),
        detail:
          tm('pdfDocxLocalDegradedDetail', { pages: imagePages.join(', ') }) +
          '\n\n' +
          tm('pdfDocxLocalOcrDetail', { pages: ocrPages.join(', ') }),
      })
    } else if (imagePages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalDegradedMsg'),
        detail: tm('pdfDocxLocalDegradedDetail', { pages: imagePages.join(', ') }),
      })
    } else if (ocrPages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalOcrMsg'),
        detail: tm('pdfDocxLocalOcrDetail', { pages: ocrPages.join(', ') }),
      })
    }
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      // structured load failures (P22): password-protected / damaged PDFs get
      // a human-readable explanation instead of the raw PDFium error string
      const detail =
        err instanceof PdfLoadError
          ? err.code === 'password-required'
            ? tm('pdfDocxLocalEncryptedDetail')
            : err.code === 'unsupported'
              ? // certificate-based or otherwise unsupported security (FPDF
                // error 5): a hard PDFium boundary — no password can open it
                // locally, so the message must NOT suggest one (P24 C)
                tm('pdfDocxLocalUnsupportedEncDetail')
              : tm('pdfDocxLocalCorruptDetail')
          : err instanceof Error
            ? err.message
            : String(err)
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfDocxFailedMsg'),
        detail,
      })
    }
  } finally {
    // the prompt window may still be open when the loop exits through cancel
    // or a non-password error thrown mid-retry
    closePdfPasswordDialog()
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

/**
 * Export as PowerPoint for pdf tabs, fully local (pdf2pptx P25): flush
 * pending edits, pick the destination, convert in-process via PDFium wasm,
 * write the file and open it in a Slides tab. No login, no credits. Shares
 * the in-flight guard with the Word exports so pdfium never runs two
 * conversions at once.
 */
async function exportPdfAsPptxLocal(): Promise<void> {
  const shellWindow = focusedShellWindow()
  const tab = focusedManager()?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfPptxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    const picked = await showSaveDialogWithMemory(
      dialog,
      shellWindow,
      {
        defaultPath: tab.filePath.replace(/\.pdf$/i, '.pptx'),
        filters: [{ name: tm('filterPpt'), extensions: ['pptx'] }],
      },
      undefined,
      tab.webContents.id,
    )
    if (picked.canceled || !picked.filePath) return
    // same stale-tab handling as the Word export (see exportPdfAsDocxLocal),
    // against the slides tab that may already show the destination file
    const staleTabId = focusedManager()?.findSlidesTabByPath(picked.filePath)
    if (staleTabId) {
      await focusedManager()?.closeTab(staleTabId)
      focusedManager()?.activateTab(tab.id)
      if (focusedManager()?.findSlidesTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    // encrypted PDFs prompt for the password (P23), looping on wrong entries;
    // null result = user cancelled the prompt → abort silently
    const pdfPath = tab.filePath
    const result = await convertPdfFileToPptxLocalWithPrompt(
      pdfPath,
      (retry) =>
        promptPdfPassword(shellWindow, {
          fileName: basename(pdfPath),
          retry,
          busy: false,
          lang: currentLang(),
          strings: {
            title: tm('pdfPwdTitle'),
            prompt: tm('pdfPwdPrompt'),
            retryPrompt: tm('pdfPwdRetryPrompt'),
            ok: tm('pdfPwdOk'),
            cancel: tm('btnCancel'),
            verifying: tm('pdfPwdVerifying'),
            label: tm('pdfPwdLabel'),
            placeholder: tm('pdfPwdPlaceholder'),
            show: tm('pdfPwdShow'),
            hide: tm('pdfPwdHide'),
          },
        }),
      (page, total) => {
        if (shellWindow && !shellWindow.isDestroyed() && total > 0) {
          shellWindow.setProgressBar(page / total)
        }
      },
    )
    if (result === null) return
    await atomicWriteFile(picked.filePath, result.pptx)

    // degrade transparency (same split as the Word export): whole scan vs
    // individual image-fallback pages
    const imagePages = result.pageResults.filter((r) => r.status !== 'ok').map((r) => r.page)
    if (result.scannedDocument) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalScannedMsg'),
        detail: tm('pdfPptxLocalScannedDetail'),
      })
    } else if (imagePages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalDegradedMsg'),
        detail: tm('pdfDocxLocalDegradedDetail', { pages: imagePages.join(', ') }),
      })
    }
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      // structured load failures (P22): same explanations as the Word export
      const detail =
        err instanceof PdfLoadError
          ? err.code === 'password-required'
            ? tm('pdfDocxLocalEncryptedDetail')
            : err.code === 'unsupported'
              ? tm('pdfDocxLocalUnsupportedEncDetail')
              : tm('pdfDocxLocalCorruptDetail')
          : err instanceof Error
            ? err.message
            : String(err)
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfPptxFailedMsg'),
        detail,
      })
    }
  } finally {
    closePdfPasswordDialog()
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

/**
 * Export as Excel for pdf tabs, fully local (pdf2xlsx P26): flush pending
 * edits, pick the destination, convert in-process via PDFium wasm, write the
 * file and open it in a Sheets tab. No login, no credits. Shares the
 * in-flight guard with the Word/PowerPoint exports so pdfium never runs two
 * conversions at once.
 */
async function exportPdfAsXlsxLocal(): Promise<void> {
  const shellWindow = focusedShellWindow()
  const tab = focusedManager()?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfXlsxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    const picked = await showSaveDialogWithMemory(
      dialog,
      shellWindow,
      {
        defaultPath: tab.filePath.replace(/\.pdf$/i, '.xlsx'),
        filters: [{ name: tm('filterExcel'), extensions: ['xlsx'] }],
      },
      undefined,
      tab.webContents.id,
    )
    if (picked.canceled || !picked.filePath) return
    // same stale-tab handling as the Word export (see exportPdfAsDocxLocal),
    // against the sheets tab that may already show the destination file
    const staleTabId = focusedManager()?.findSheetsTabByPath(picked.filePath)
    if (staleTabId) {
      await focusedManager()?.closeTab(staleTabId)
      focusedManager()?.activateTab(tab.id)
      if (focusedManager()?.findSheetsTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    // encrypted PDFs prompt for the password (P23), looping on wrong entries;
    // null result = user cancelled the prompt → abort silently
    const pdfPath = tab.filePath
    const result = await convertPdfFileToXlsxLocalWithPrompt(
      pdfPath,
      (retry) =>
        promptPdfPassword(shellWindow, {
          fileName: basename(pdfPath),
          retry,
          busy: false,
          lang: currentLang(),
          strings: {
            title: tm('pdfPwdTitle'),
            prompt: tm('pdfPwdPrompt'),
            retryPrompt: tm('pdfPwdRetryPrompt'),
            ok: tm('pdfPwdOk'),
            cancel: tm('btnCancel'),
            verifying: tm('pdfPwdVerifying'),
            label: tm('pdfPwdLabel'),
            placeholder: tm('pdfPwdPlaceholder'),
            show: tm('pdfPwdShow'),
            hide: tm('pdfPwdHide'),
          },
        }),
      (page, total) => {
        if (shellWindow && !shellWindow.isDestroyed() && total > 0) {
          shellWindow.setProgressBar(page / total)
        }
      },
    )
    if (result === null) return
    await atomicWriteFile(picked.filePath, result.xlsx)

    // degrade transparency: pages that could not become cells got a notice
    // row on their worksheet instead of an image (a spreadsheet has none)
    const noticePages = result.pageResults.filter((r) => r.status !== 'ok').map((r) => r.page)
    if (result.scannedDocument) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalScannedMsg'),
        detail: tm('pdfXlsxLocalScannedDetail'),
      })
    } else if (noticePages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfXlsxLocalSkippedMsg'),
        detail: tm('pdfXlsxLocalSkippedDetail', { pages: noticePages.join(', ') }),
      })
    }
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      // structured load failures (P22): same explanations as the Word export
      const detail =
        err instanceof PdfLoadError
          ? err.code === 'password-required'
            ? tm('pdfDocxLocalEncryptedDetail')
            : err.code === 'unsupported'
              ? tm('pdfDocxLocalUnsupportedEncDetail')
              : tm('pdfDocxLocalCorruptDetail')
          : err instanceof Error
            ? err.message
            : String(err)
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfXlsxFailedMsg'),
        detail,
      })
    }
  } finally {
    closePdfPasswordDialog()
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

// The pdf renderer's converter dropdown funnels into the same local conversion
// flows as the File menu items (dialogs, password prompt, in-flight guard included)
ipcMain.handle(PDF_CHANNELS.convertOffice, async (e, format: unknown) => {
  // only the sender's own active pdf tab may trigger a conversion (its file
  // is the source) — resolve the tab's window, not the focused one
  if (managerForWebContents(e.sender.id)?.activePdfTab()?.webContents.id !== e.sender.id) return
  if (format === 'docx') await exportPdfAsDocxLocal()
  else if (format === 'xlsx') await exportPdfAsXlsxLocal()
  else if (format === 'pptx') await exportPdfAsPptxLocal()
})

function openThirdPartyNotices(): Promise<string> {
  const path = app.isPackaged
    ? join(process.resourcesPath, 'THIRD-PARTY-NOTICES.txt')
    : join(app.getAppPath(), 'build', 'THIRD-PARTY-NOTICES.txt')
  return shell.openPath(path)
}

/** every module's File menu gets the New submenu (parity with Home, FIRST —
 *  Office convention, matching the shell's own File menus) and a way back to
 *  the launcher */
function installBackToHomeItems(): void {
  const backToHomeItem: MenuItemConstructorOptions = {
    label: tm('backToHome'),
    accelerator: 'Shift+CmdOrCtrl+H',
    click: () => focusedManager()?.openHomeTab(),
  }
  setDocsFileMenuHeadItems([newFileSubMenu()])
  setSheetsFileMenuHeadItems([newFileSubMenu()])
  setSlidesFileMenuHeadItems([newFileSubMenu()])
  setDocsExtraFileMenuItems([backToHomeItem])
  setSheetsExtraFileMenuItems([backToHomeItem])
  setSlidesExtraFileMenuItems([backToHomeItem])
}

function installDockMenu(): void {
  if (process.platform !== 'darwin') return
  app.dock?.setMenu(
    Menu.buildFromTemplate([
      { label: tm('menuHome'), click: () => focusedManager()?.openHomeTab() },
      {
        label: tm('menuNewDoc'),
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        click: () => void newSheetTab(),
      },
      { label: tm('menuNewSlide'), click: () => newSlideTab() },
      { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
      { label: tm('menuNewHtml'), click: () => newHtmlTab() },
      { label: tm('menuNewPdf'), click: () => void newPdfTab() },
    ]),
  )
}

// On mainland-China networks the main process's Node fetch (undici) bypasses the system proxy,
// so direct calls to overseas LLM/image-search APIs time out or get region-blocked (403).
// Prefer proxy env vars (terminal launch); a packaged app launched from Finder inherits no shell
// env vars, so fall back to the system HTTP proxy. The renderer uses Chromium's system proxy and
// is unaffected. Same bootstrap as slides-main startSlidesStandalone.
async function installMainProcessProxy(): Promise<void> {
  let proxyUrl = [
    process.env.HTTPS_PROXY,
    process.env.https_proxy,
    process.env.HTTP_PROXY,
    process.env.http_proxy,
    process.env.ALL_PROXY,
    process.env.all_proxy,
  ].find((v) => v && /^https?:\/\//.test(v))
  if (!proxyUrl) {
    try {
      // PAC/rule proxies answer per-host: probe a host the LLM/image-search
      // APIs live behind
      const resolved = await session.defaultSession.resolveProxy('https://api.openai.com/')
      const m = /PROXY\s+([^;\s]+)/.exec(resolved)
      if (m) proxyUrl = `http://${m[1]}`
    } catch {
      /* no system proxy */
    }
  }
  if (!proxyUrl) return
  try {
    const { ProxyAgent, setGlobalDispatcher } = await import('undici')
    setGlobalDispatcher(new ProxyAgent(proxyUrl))
    // strip user:pass credentials before logging
    console.log('[proxy] main-process fetch via', proxyUrl.replace(/\/\/[^@/]*@/, '//***@'))
  } catch (e) {
    console.warn('[proxy] failed to set ProxyAgent:', e)
  }
}

// ---- lifecycle (the shell is the only owner) ----

let pendingLaunchPath =
  supportedFileIn(process.argv) ?? unsupportedFileIn(process.argv) ?? missingFileIn(process.argv)

// show() does not un-minimize, and on macOS ⌘W destroys the shell window while the
// app keeps running — either way a file opened from Finder would land out of sight.
function revealShellWindow(): void {
  if (shellEntries().length === 0) createShellWindow({ primary: true })
  const win = focusedShellWindow()
  if (win?.isMinimized()) win.restore()
  win?.show()
  win?.focus()
}

// On macOS a file opened from Finder is not in argv; it arrives via the open-file event (before ready).
// If another instance already holds the lock, this process exits, and the path must ride along in
// the lock request's additionalData to the surviving instance — so the lock request is deferred
// until ready, after the path is known.
app.on('open-file', (event, filePath) => {
  event.preventDefault()
  if (!app.isReady()) {
    pendingLaunchPath = filePath
    return
  }
  revealShellWindow()
  if (!openDocumentPath(filePath)) {
    // an intended open that produced no tab must explain itself (BUG-1655)
    reportOpenFailure(filePath, openFailureDeps)
    focusedManager()?.openHomeTab()
  }
})

app.on('second-instance', (_event, argv, _cwd, additionalData) => {
  const file =
    supportedFileIn(argv) ??
    unsupportedFileIn(argv) ??
    missingFileIn(argv) ??
    (additionalData as { launchPath?: string } | null)?.launchPath
  revealShellWindow()
  if (!file) {
    focusedManager()?.openHomeTab()
    return
  }
  if (!openDocumentPath(file)) {
    // an intended open that produced no tab must explain itself (BUG-1655)
    reportOpenFailure(file, openFailureDeps)
    focusedManager()?.openHomeTab()
  }
})

installNavigationGuard(app)
installContextMenu(app, () => contextMenuLabels(currentLang()))
registerAiIpc()
registerProjectIpc()
registerDocsIpc()
registerHomeIpc()
registerTabsIpc()
registerDroppedFilesIpc()
installShellModuleHooks()

// sheets' project:resolveChat goes through the handler registered by docs-main; the sessionId reverse lookup hooks in here
setSessionPathResolver(resolveSheetsSessionPath)

/** Dev-only pid marker for the takeover below; scoped to userData like the lock itself. */
const devPidFile = () => join(app.getPath('userData'), 'dev-instance.pid')

app.whenReady().then(async () => {
  const lockData = () => (pendingLaunchPath ? { launchPath: pendingLaunchPath } : {})
  let hasLock = app.requestSingleInstanceLock(lockData())
  if (!hasLock && !app.isPackaged) {
    // Dev watch restart: electron-vite SIGTERMs the previous instance and spawns this
    // one immediately. Chromium turns that SIGTERM into a graceful quit (Node's
    // process.on('SIGTERM') never fires in the main process), and the quit can wedge
    // in the close-confirmation flow — the zombie then keeps the single-instance lock,
    // this instance quits, and electron-vite's on-close handler exits with it, killing
    // the renderer dev server (blank shell window until a manual dev restart).
    // The previous instance is doomed either way: kill it and take over the lock.
    try {
      const oldPid = Number(readFileSync(devPidFile(), 'utf-8').trim())
      if (Number.isFinite(oldPid) && oldPid > 0 && oldPid !== process.pid) {
        // pid-recycling guard: only kill if that pid is still an Electron process
        const cmd = execSync(`ps -o command= -p ${oldPid}`).toString()
        if (cmd.includes('Electron')) process.kill(oldPid, 'SIGKILL')
      }
    } catch {
      // no previous instance recorded / already gone (ps exits non-zero)
    }
    for (let i = 0; i < 20 && !hasLock; i++) {
      await new Promise((r) => setTimeout(r, 150))
      hasLock = app.requestSingleInstanceLock(lockData())
    }
  }
  if (!hasLock) {
    app.quit()
    return
  }
  if (!app.isPackaged) {
    try {
      writeFileSync(devPidFile(), String(process.pid))
    } catch {
      // best-effort: without the marker the next restart just retries the lock
    }
  }

  void installMainProcessProxy()
  app.setAccessibilitySupportEnabled(true)
  // Settle the shared uiLang from saved settings BEFORE any tab renderer can
  // ask 'app:get-language': the editor handlers return the i18n module's
  // mutable lang, whose 'zh' default otherwise wins the race for whichever
  // tab loads first (e.g. sheets booting in Chinese while docs shows English).
  currentLang()
  // native menus/dialogs/scrollbars follow the persisted theme from first paint
  nativeTheme.themeSource = currentTheme()
  // stamp the star-prompt install-age clock on the first launch carrying the feature,
  // and detect upgrade launches (version changed since the previous run)
  try {
    const settings = readAppSettings(APP_SETTINGS_PATH())
    const starState = readStarPrompt()
    const stamped = withFirstRun(starState, Date.now())
    if (stamped !== starState) writeStarPrompt(stamped)

    const prevVersion =
      typeof settings[LAST_RUN_VERSION_KEY] === 'string'
        ? (settings[LAST_RUN_VERSION_KEY] as string)
        : null
    const currentVersion = app.getVersion()
    upgradeStarPromptPending = isUpgradeLaunch(
      prevVersion,
      currentVersion,
      settings.onboardingSeen === true,
    )
    if (prevVersion !== currentVersion)
      writeAppSetting(APP_SETTINGS_PATH(), LAST_RUN_VERSION_KEY, currentVersion)
  } catch {
    // settings write failures must never block startup
  }
  startSheetsCaptureServer()
  // live bridge for external agents (Airy Copilot): on by default; off via
  // Settings → General or AIRY_DISABLE_BRIDGE=1 (agents keep working unless
  // the user explicitly opts out, so the default stays on)
  if (liveBridgeEnabled()) {
    void startLiveBridge().catch((err: unknown) => {
      console.error('bridge server failed to start:', err)
    })
  }
  // In-app updater backed by the fork's GitHub Releases (see
  // src/main/updater/): inactive in dev; macOS shows a manual-download
  // dialog from Help > Check for Updates (no feed contact). The first check
  // is deferred inside initUpdater so startup never waits on the network.
  // Must run before createShellWindow: the menu builders read its state.
  initUpdater({
    isPackaged: app.isPackaged,
    getWindow: () => focusedShellWindow(),
    getLabels: () => ({
      check: tm('menuCheckUpdates'),
      checking: tm('updStatusChecking'),
      downloading: tm('updStatusDownloading'),
      ready: tm('updStatusReady'),
      upToDate: tm('updStatusUpToDate'),
      failed: tm('updStatusFailed'),
    }),
    onStatusChange: () => applyMenuFor(currentMenuKind),
  })
  createShellWindow({ primary: true })
  // deferred to ready: labels need currentLang(), which reads app.getLocale()
  installBackToHomeItems()
  installDockMenu()

  const restoredTabs = restorePreviousSession()
  if (!pendingLaunchPath || !openDocumentPath(pendingLaunchPath)) {
    // a launch path that cannot be read explains itself (BUG-1655) instead of
    // quietly doing nothing
    if (pendingLaunchPath) reportOpenFailure(pendingLaunchPath, openFailureDeps)
    // nothing to open and no session to fall back on → Home
    if (restoredTabs === 0) focusedManager()?.openHomeTab()
  }
  pendingLaunchPath = null
  // staged untitled files nothing reopened (crash leftovers whose session was
  // not restored — restore disabled or pruned) are dead scratch: clean them
  purgeOrphanStagedFiles()

  app.on('activate', () => {
    if (shellEntries().length === 0) createShellWindow({ primary: true })
  })
})

app.on('window-all-closed', () => {
  // an updater "Install now" rides the quit's close flow (BUG-411): every
  // dirty guard has passed by here, so hand over to the installer before the
  // final quit — never while windows that could still Cancel are alive
  flushPendingInstall()
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', () => {
  quitFlow.begin()
  // BUG-408: the tab-change session save is debounced by 800 ms, so a crash
  // or hard kill once the quit has started (a renderer dying under an open
  // dirty prompt, SIGKILL mid-shutdown) used to lose the last tab changes.
  // Land a fresh full snapshot synchronously right away — every window is
  // still registered here, so this write can only ever WIDEN the persisted
  // set, never shrink it like a mid-quit write would. The quit-time
  // persist-once still writes again at the first confirmed close (fresher,
  // minus the windows that confirmed); neither write is counted as the other.
  persistSessionState(false)
  // no close prompt may fall through to "Save" during shutdown, the sidecar
  // and pdf workers stop, and the live bridge closes its socket — the whole
  // set rolls back if a cancelled dirty-guard aborts the quit (BUG-1216)
  shutdownEffects.arm()
})

app.on('will-quit', () => {
  // BUG-1307: the quit snapshot may still be missing when the quit completes
  // — a write failure on the LAST confirmed close has no successor close the
  // BUG-1224 re-arm could retry into. One final write here (the same full
  // snapshot before-quit makes); only a real write failure logs, so the
  // empty-registry flush (nothing left to serialize) stays silent.
  if (!quitFlow.snapshotWritePending()) return
  if (!persistSessionState(false)) {
    console.error(
      '[shell] quit-time session snapshot write failed again at will-quit; the next ' +
        'launch restores the stale snapshot and may resurrect windows closed before ' +
        'the failure',
    )
  }
})

// Electron glue for the in-app updater: wires electron-updater's autoUpdater
// (GitHub Releases feed of the fork) and Electron's dialog/Notification into
// the pure controller (core.ts). This is the only updater file that imports
// Electron. The shell menu integration (index.ts) renders the status via
// updaterMenuItems() and re-runs its menu builder on every status change —
// no dedicated renderer window, following the bridge/ glue split.
import { Notification, app, dialog } from 'electron'
import type { BrowserWindow, MenuItemConstructorOptions } from 'electron'
import { autoUpdater } from 'electron-updater'

import { openHelpUrl } from '@airy-office/electron-utils'

import {
  GITHUB_OWNER,
  GITHUB_REPO,
  UpdaterController,
  createPendingInstaller,
  detectUpdatePolicy,
  type UpdaterClient,
  type UpdaterUi,
} from './core'
import { updaterText } from './strings'

/** menu labels for the current status; injected by index.ts (owns the tMain table) */
export interface UpdaterMenuLabels {
  check: string
  checking: string
  downloading: string // may carry a {percent} placeholder
  ready: string
  upToDate: string
  failed: string
}

export interface UpdaterGlueOptions {
  /** dev runs never check the feed (no app-update.yml, dev version) */
  isPackaged: boolean
  /** parent for the dialogs (menus/dialogs stay attached to the shell window) */
  getWindow: () => BrowserWindow | null
  /** menu labels in the current UI language, re-read on every menu build */
  getLabels: () => UpdaterMenuLabels
  /** status fan-out: the shell rebuilds its application menu here */
  onStatusChange?: () => void
  /** override for tests of the glue layer itself */
  policy?: ReturnType<typeof detectUpdatePolicy>
  startDelayMs?: number
}

let controller: UpdaterController | null = null
let labelsProvider: (() => UpdaterMenuLabels) | null = null

/**
 * "Install now" goes through the app's own close flow instead of
 * electron-updater's quitAndInstall directly (BUG-411): the request arms the
 * gate and starts a plain app.quit(), every window walks its dirty-tab guard,
 * and the install fires only from the flush points once no window can object.
 * A cancelled guard disarms it (cancelPendingInstall from the shell's
 * abortAppQuit) — the staged update then installs on a later natural quit.
 */
const pendingInstaller = createPendingInstaller()

/** flush point: all windows are gone — hand over to the installer */
export function flushPendingInstall(): void {
  if (pendingInstaller.flush()) autoUpdater.quitAndInstall(true, true)
}

/** a dirty-guard Cancel aborted the quit the install was riding on */
export function cancelPendingInstall(): void {
  pendingInstaller.cancel()
}

/**
 * Configure the updater. Inactive (no menu item, no checks) when the app is
 * unpacked or the platform policy is 'off' (unknown platforms). macOS uses
 * the 'manual' policy: the menu item shows a dialog linking to the releases
 * page and no feed is ever contacted. Called once from app.whenReady; never
 * blocks startup — the first check is deferred by the controller (core.ts
 * DEFAULT_START_DELAY_MS).
 */
export function initUpdater(options: UpdaterGlueOptions): void {
  if (controller) return
  labelsProvider = options.getLabels
  const policy =
    options.policy !== undefined ? options.policy : detectUpdatePolicy(process.platform)
  if (!options.isPackaged || policy === 'off') {
    controller = new UpdaterController({
      policy: 'off',
      client: neverClient(),
      ui: neverUi(),
      log,
    })
    return
  }
  if (policy === 'manual') {
    controller = new UpdaterController({
      policy,
      client: neverClient(),
      ui: manualUi(options.getWindow),
      log,
      onStatusChange: options.onStatusChange,
    })
    return
  }

  // Feed is pinned to the fork here even though electron-builder also bakes
  // the same publish config into resources/app-update.yml: an explicit
  // setFeedURL keeps dev tooling and future repackaging from silently
  // pointing installs at the wrong repository.
  autoUpdater.setFeedURL({ provider: 'github', owner: GITHUB_OWNER, repo: GITHUB_REPO })
  autoUpdater.autoDownload = false
  // a user picking "Install on quit" still gets the update on normal exit
  autoUpdater.autoInstallOnAppQuit = true

  controller = new UpdaterController({
    policy,
    client: autoUpdaterClient(),
    ui: dialogUi(options.getWindow),
    log,
    onStatusChange: options.onStatusChange,
    startDelayMs: options.startDelayMs,
  })
  controller.start()
  // second flush point (the shell's window-all-closed handler is the first):
  // a quit with zero windows left never emits window-all-closed, and the
  // natural-quit path reaches this with nothing pending — a cheap no-op
  app.on('will-quit', () => {
    flushPendingInstall()
  })
}

function log(message: string, error?: unknown): void {
  // expected conditions (offline machine, GitHub rate limits) stay stderr-only
  console.error('[updater]', message, error ?? '')
}

function autoUpdaterClient(): UpdaterClient {
  return {
    async checkForUpdates() {
      const result = await autoUpdater.checkForUpdates()
      const version = result?.updateInfo?.version ?? null
      return { available: result?.isUpdateAvailable === true && version !== null, version }
    },
    downloadUpdate: () => autoUpdater.downloadUpdate(),
    // BUG-411: request the install through the app's own close flow — see
    // pendingInstaller above; the flush points call quitAndInstall for real
    quitAndInstall: () => {
      pendingInstaller.request()
      app.quit()
    },
    onDownloadProgress: (callback) => {
      autoUpdater.on('download-progress', (progress) => callback(progress.percent))
    },
    onError: (callback) => {
      autoUpdater.on('error', (error) => callback(error))
    },
  }
}

function dialogUi(getWindow: () => BrowserWindow | null): UpdaterUi {
  const show = async (options: Electron.MessageBoxOptions) => {
    const win = getWindow()
    return win && !win.isDestroyed()
      ? dialog.showMessageBox(win, options)
      : dialog.showMessageBox(options)
  }
  return {
    async offerUpdate({ version, notesUrl }) {
      const { response } = await show({
        type: 'info',
        title: updaterText('updAvailableTitle'),
        message: updaterText('updAvailableBody', { version }),
        // native dialogs cannot hyperlink: the dedicated button opens the
        // notes page instead, and the URL stays visible as detail text
        detail: notesUrl,
        buttons: [
          updaterText('updDownloadButton'),
          updaterText('updNotesButton'),
          updaterText('updLaterButton'),
        ],
        defaultId: 0,
        cancelId: 2,
        noLink: true,
      })
      // single gate for every shell.openExternal (see github-menu's openHelpUrl)
      if (response === 1) void openHelpUrl(notesUrl)
      return response === 0 ? 'download' : 'later'
    },
    async offerInstall({ version }) {
      const { response } = await show({
        type: 'info',
        title: updaterText('updReadyTitle'),
        message: updaterText('updReadyBody', { version }),
        buttons: [updaterText('updInstallNow'), updaterText('updInstallOnQuit')],
        defaultId: 0,
        cancelId: 1,
        noLink: true,
      })
      return response === 0 ? 'now' : 'on-quit'
    },
    notifyAboutRelease({ version, releasesUrl }) {
      const notification = new Notification({
        title: updaterText('updNotifyTitle'),
        body: updaterText('updNotifyBody', { version }),
      })
      notification.on('click', () => void openHelpUrl(releasesUrl))
      notification.show()
    },
    // the dialog UI serves the in-app/notify-only policies; 'manual' gets its
    // own UI below, so this entry is never reached
    notifyManualUpdate: () => {},
  }
}

/** macOS 'manual' policy UI: a dialog linking to the releases page (the feed
 *  is never queried, so there is no version to report — just the way out) */
function manualUi(getWindow: () => BrowserWindow | null): UpdaterUi {
  return {
    async offerUpdate() {
      return 'later'
    },
    async offerInstall() {
      return 'on-quit'
    },
    notifyAboutRelease: () => {},
    notifyManualUpdate({ releasesUrl }) {
      const options: Electron.MessageBoxOptions = {
        type: 'info',
        title: updaterText('updManualTitle'),
        message: updaterText('updManualBody'),
        // native dialogs cannot hyperlink: the dedicated button opens the
        // releases page, and the URL stays visible as detail text
        detail: releasesUrl,
        buttons: [updaterText('updManualButton'), updaterText('updLaterButton')],
        defaultId: 0,
        cancelId: 1,
        noLink: true,
      }
      const win = getWindow()
      const shown =
        win && !win.isDestroyed()
          ? dialog.showMessageBox(win, options)
          : dialog.showMessageBox(options)
      void shown.then(({ response }) => {
        // single gate for every shell.openExternal (see github-menu's openHelpUrl)
        if (response === 0) void openHelpUrl(releasesUrl)
      })
    },
  }
}

/** inert implementations for the inactive glue (dev runs / macOS) */
function neverClient(): UpdaterClient {
  return {
    async checkForUpdates() {
      return { available: false, version: null }
    },
    async downloadUpdate() {
      return undefined
    },
    quitAndInstall: () => {},
    onDownloadProgress: () => {},
    onError: () => {},
  }
}

function neverUi(): UpdaterUi {
  return {
    async offerUpdate() {
      return 'later'
    },
    async offerInstall() {
      return 'on-quit'
    },
    notifyAboutRelease: () => {},
    notifyManualUpdate: () => {},
  }
}

/**
 * Menu items for the shell's Help submenu: empty while inactive. The item
 * label mirrors the live status (check → checking → downloading x% → ready);
 * non-idle statuses disable the item so a click cannot start a second flow.
 */
export function updaterMenuItems(): MenuItemConstructorOptions[] {
  const ctl = controller
  const active = ctl !== null && ctl.active
  if (!active || !labelsProvider) return []
  const labels = labelsProvider()
  const status = ctl.status
  const item: MenuItemConstructorOptions =
    status.phase === 'idle'
      ? { label: labels.check, click: () => ctl.checkNow() }
      : status.phase === 'checking'
        ? { label: labels.checking, enabled: false }
        : status.phase === 'downloading'
          ? {
              label: labels.downloading.replace('{percent}', String(status.percent)),
              enabled: false,
            }
          : status.phase === 'ready'
            ? { label: labels.ready, enabled: false }
            : status.phase === 'up-to-date'
              ? { label: labels.upToDate, enabled: false }
              : { label: labels.failed, enabled: false }
  return [item]
}

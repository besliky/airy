// Pure updater state machine (no Electron imports): platform policy, the
// check -> available -> downloading -> ready flow, and the transient status
// the shell's menu item renders. Every side effect (electron-updater calls,
// dialogs, notifications, menu rebuilds) is injected, so the whole machine
// is covered by tests/updater.test.ts against mocks (see bridge/ for the
// same pure-core / thin-glue split).

/** the fork's GitHub Releases feed (electron-builder publish config + manual URLs) */
export const GITHUB_OWNER = 'besliky'
export const GITHUB_REPO = 'airy'

/** release-notes page for one version (tags follow the vX.Y.Z convention) */
export function releaseNotesUrl(version: string): string {
  return `https://github.com/${GITHUB_OWNER}/${GITHUB_REPO}/releases/tag/v${version}`
}

/** where notify-only installs (deb) are pointed — always the newest release */
export const RELEASES_LATEST_URL = `https://github.com/${GITHUB_OWNER}/${GITHUB_REPO}/releases/latest`

/**
 * What this install can do with an update:
 *  - 'in-app'      Windows NSIS and Linux AppImage: checkForUpdates +
 *                  downloadUpdate + quitAndInstall, all inside the app.
 *  - 'notify-only' Linux deb installs: electron-updater cannot self-update a
 *                  deb package, so an available update only raises a
 *                  notification linking to the releases page.
 *  - 'manual'      macOS: unsigned builds cannot auto-update (Squirrel/
 *                  electron-updater would need signed dmg + latest-mac.yml
 *                  trust); the menu item shows a dialog linking to the
 *                  releases page, and no feed check ever runs automatically.
 *  - 'off'         dev runs and unknown platforms: inactive.
 */
export type UpdatePolicy = 'in-app' | 'notify-only' | 'manual' | 'off'

/**
 * Platform detection, injected so tests stay deterministic. On Linux the
 * AppImage runtime exports APPIMAGE pointing at the running .AppImage file;
 * its absence means the app came from the deb (or another package), which
 * electron-updater cannot self-update.
 */
export function detectUpdatePolicy(
  platform: NodeJS.Platform,
  env: { APPIMAGE?: string } = process.env,
): UpdatePolicy {
  if (platform === 'win32') return 'in-app'
  if (platform === 'linux') return env.APPIMAGE ? 'in-app' : 'notify-only'
  if (platform === 'darwin') return 'manual'
  return 'off'
}

/** menu-facing status; the transient phases ('up-to-date', 'failed') revert to idle */
export type UpdaterStatus =
  | { phase: 'idle' }
  | { phase: 'checking' }
  | { phase: 'downloading'; percent: number }
  | { phase: 'ready' }
  | { phase: 'up-to-date' }
  | { phase: 'failed' }

/** autoUpdater-like surface the controller drives (electron-updater in prod, mocks in tests) */
export interface UpdaterClient {
  /** resolves when the check settles; available carries the candidate version */
  checkForUpdates(): Promise<{ available: boolean; version: string | null }>
  /** background download to the update cache */
  downloadUpdate(): Promise<unknown>
  /** quit, apply the staged update and relaunch (NSIS / AppImage) */
  quitAndInstall(): void
  /** progress ticks while downloadUpdate runs (percent 0..100) */
  onDownloadProgress(callback: (percent: number) => void): void
  /** spurious autoUpdater errors outside our awaited calls (logged, never dialoged) */
  onError(callback: (error: unknown) => void): void
}

/** user-facing surfaces the controller drives (Electron dialog/Notification in prod) */
export interface UpdaterUi {
  /**
   * Offer an in-app update. The "Release Notes" button opens notesUrl inside
   * the implementation (native dialogs cannot hyperlink) and resolves 'later'.
   */
  offerUpdate(info: { version: string; notesUrl: string }): Promise<'download' | 'later'>
  /** download finished: install right now (restarts the app) or on quit */
  offerInstall(info: { version: string }): Promise<'now' | 'on-quit'>
  /** notify-only policy: click-through notice pointing at the releases page */
  notifyAboutRelease(info: { version: string; releasesUrl: string }): void
  /** manual policy (macOS): dialog pointing at the releases page — no version
   *  is known because the feed is never queried */
  notifyManualUpdate(info: { releasesUrl: string }): void
}

export interface UpdaterControllerOptions {
  client: UpdaterClient
  ui: UpdaterUi
  policy: UpdatePolicy
  /** diagnostic sink; network failures are expected and must stay silent otherwise */
  log?: (message: string, error?: unknown) => void
  /** status fan-out for the menu item label */
  onStatusChange?: (status: UpdaterStatus) => void
  /** deferred startup check delay (see start) */
  startDelayMs?: number
  /** how long the transient up-to-date/failed statuses stay in the menu */
  transientStatusMs?: number
}

export const DEFAULT_START_DELAY_MS = 10_000
export const DEFAULT_TRANSIENT_STATUS_MS = 5_000

/**
 * Orchestrates one update flow at a time. All failures — network, GitHub API
 * limits, download aborts — funnel into the log plus a transient 'failed'
 * menu status; an error dialog is never shown (a background check that
 * cannot reach the feed is normal, not something to bother the user with).
 */
export class UpdaterController {
  readonly policy: UpdatePolicy
  private readonly client: UpdaterClient
  private readonly ui: UpdaterUi
  private readonly log: (message: string, error?: unknown) => void
  private readonly onStatusChange?: (status: UpdaterStatus) => void
  private readonly transientStatusMs: number
  private statusValue: UpdaterStatus = { phase: 'idle' }
  private transientTimer: NodeJS.Timeout | null = null
  private started = false

  constructor(options: UpdaterControllerOptions) {
    this.client = options.client
    this.ui = options.ui
    this.policy = options.policy
    this.log = options.log ?? (() => {})
    this.onStatusChange = options.onStatusChange
    this.transientStatusMs = options.transientStatusMs ?? DEFAULT_TRANSIENT_STATUS_MS
    this.client.onDownloadProgress((percent) => this.handleProgress(percent))
    this.client.onError((error) => this.handleAsyncError(error))
  }

  get active(): boolean {
    return this.policy !== 'off'
  }

  get status(): UpdaterStatus {
    return this.statusValue
  }

  /** true while a flow occupies the menu status (re-checks are ignored) */
  get busy(): boolean {
    const phase = this.statusValue.phase
    return phase === 'checking' || phase === 'downloading' || phase === 'ready'
  }

  /**
   * Schedule the deferred startup check: a few seconds after whenReady, so
   * the update check never competes with app boot. No-op for the 'off'
   * policy (dev) and 'manual' (macOS: no automatic feed contact — the check
   * only ever happens from the menu) — call start() unconditionally from
   * the glue.
   */
  start(startDelayMs: number = DEFAULT_START_DELAY_MS): void {
    if (this.started || !this.active || this.policy === 'manual') return
    this.started = true
    setTimeout(() => {
      this.checkNow()
    }, startDelayMs)
  }

  /** manual check from the menu; also the auto-check entry point */
  checkNow(): void {
    if (!this.active || this.busy) return
    // macOS: the menu item is a documented stub — show the manual-download
    // dialog instead of contacting the feed
    if (this.policy === 'manual') {
      this.ui.notifyManualUpdate({ releasesUrl: RELEASES_LATEST_URL })
      return
    }
    this.setStatus({ phase: 'checking' })
    this.client
      .checkForUpdates()
      .then((result) => this.handleCheckResult(result))
      .catch((error: unknown) => this.handleFailure('update check failed', error))
  }

  /** download progress tick; ignored outside an in-flight download */
  handleProgress(percent: number): void {
    if (this.statusValue.phase !== 'downloading') return
    this.setStatus({ phase: 'downloading', percent: clampPercent(percent) })
  }

  /**
   * autoUpdater 'error' event with no awaited call to surface it (e.g. a
   * background check racing a manual one): logged; while a download is in
   * flight it also parks the flow, because the downloadUpdate promise
   * rejects through this same event.
   */
  handleAsyncError(error: unknown): void {
    this.log('autoUpdater error', error)
    if (this.statusValue.phase === 'downloading') {
      this.handleFailure('update download failed', error)
    }
  }

  private handleCheckResult(result: { available: boolean; version: string | null }): void {
    if (!result.available || !result.version) {
      this.setTransientStatus({ phase: 'up-to-date' })
      return
    }
    if (this.policy === 'notify-only') {
      this.ui.notifyAboutRelease({ version: result.version, releasesUrl: RELEASES_LATEST_URL })
      this.setStatus({ phase: 'idle' })
      return
    }
    void this.offerInAppUpdate(result.version)
  }

  private async offerInAppUpdate(version: string): Promise<void> {
    const choice = await this.ui.offerUpdate({ version, notesUrl: releaseNotesUrl(version) })
    if (choice !== 'download') {
      this.setStatus({ phase: 'idle' })
      return
    }
    this.setStatus({ phase: 'downloading', percent: 0 })
    try {
      await this.client.downloadUpdate()
      this.setStatus({ phase: 'ready' })
      // 'now' restarts into the installer; 'on-quit' leans on
      // autoInstallOnAppQuit, which the glue leaves enabled
      if ((await this.ui.offerInstall({ version })) === 'now') this.client.quitAndInstall()
    } catch (error) {
      this.handleFailure('update download failed', error)
    }
  }

  private handleFailure(message: string, error: unknown): void {
    this.log(message, error)
    this.setTransientStatus({ phase: 'failed' })
  }

  private setStatus(status: UpdaterStatus): void {
    if (this.transientTimer) {
      clearTimeout(this.transientTimer)
      this.transientTimer = null
    }
    this.statusValue = status
    this.onStatusChange?.(status)
  }

  private setTransientStatus(status: UpdaterStatus): void {
    this.setStatus(status)
    this.transientTimer = setTimeout(() => {
      this.transientTimer = null
      if (this.statusValue === status) this.setStatus({ phase: 'idle' })
    }, this.transientStatusMs)
  }
}

function clampPercent(percent: number): number {
  if (!Number.isFinite(percent) || percent < 0) return 0
  return Math.min(100, Math.round(percent))
}

/**
 * Gate between "install now" and actually handing over to the installer
 * (BUG-411). electron-updater's quitAndInstall spawns the NSIS/AppImage
 * installer BEFORE its own app.quit(), so calling it directly while windows
 * are still open let a dirty-guard Cancel abort the quit with the installer
 * already waiting outside — unsaved work on one side, a half-started update
 * on the other. The request instead rides a normal app.quit(): every window
 * walks its close guard, and the flush (wired to window-all-closed / will-quit)
 * calls quitAndInstall only once no window is left to object. A Cancel after
 * the request disarms it — the staged update still installs on a later
 * natural quit via autoInstallOnAppQuit.
 */
export interface PendingInstaller {
  get pending(): boolean
  request(): void
  /** a dirty-guard Cancel aborted the quit the install was riding on */
  cancel(): void
  /** true exactly once, when a request survived up to the flush point */
  flush(): boolean
}

export function createPendingInstaller(): PendingInstaller {
  let pending = false
  return {
    get pending() {
      return pending
    },
    request() {
      pending = true
    },
    cancel() {
      pending = false
    },
    flush() {
      if (!pending) return false
      pending = false
      return true
    },
  }
}

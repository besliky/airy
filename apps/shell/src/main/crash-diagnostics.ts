import { mkdirSync } from 'node:fs'
import { join } from 'node:path'
import { app, crashReporter } from 'electron'
import { createMainLog, teeConsoleInto, type MainLog } from './main-log'

/**
 * Local-only crash diagnostics for the shell main process (UX-1628): the
 * complaint "the app closed itself" used to be undiagnosable because
 * crashReporter was never started (zero minidumps in the real profile with
 * the report) and main-process logs only reached stderr, which a desktop
 * launcher discards.
 *
 * Privacy contract: everything stays on the user's machine. The crash
 * reporter is started with uploadToServer:false and no submitURL, so
 * minidumps are only ever written under userData/crash-dumps; the ring log
 * lives under userData/logs. Nothing here performs network I/O.
 *
 * Failure containment: diagnostics sit on the app's crash path, so every
 * step is internally guarded — a telemetry failure must never take the app
 * (or an in-flight crash report) down with it.
 */

export const CRASH_DUMPS_DIR_NAME = 'crash-dumps'
export const LOGS_DIR_NAME = 'logs'

/** minimal structural Electron app surface (keeps the wiring unit-testable) */
interface AppEventTarget {
  setPath(name: string, path: string): void
  getPath(name: string): string
  on(event: string, listener: (...args: unknown[]) => void): unknown
}

/** minimal structural crashReporter surface (start options we rely on) */
interface CrashReporterLike {
  start(options: { submitURL?: string; uploadToServer?: boolean; compress?: boolean }): void
}

/** structured log target for the fatal-event wiring (subset of MainLog) */
type EventLog = Pick<MainLog, 'event'>

/**
 * Start the Crashpad reporter with LOCAL minidumps only: the dump directory
 * is pinned under userData (the app's 'crashDumps' path) and uploads are
 * disabled at the API level (uploadToServer:false, no submitURL given).
 * Never throws.
 */
export function startLocalCrashReporter(deps: {
  app: Pick<AppEventTarget, 'setPath' | 'getPath'>
  crashReporter: CrashReporterLike
  userDataDir: string
}): { dumpsDir: string } {
  const dumpsDir = join(deps.userDataDir, CRASH_DUMPS_DIR_NAME)
  try {
    // pin the Chromium 'crashDumps' path BEFORE the reporter starts: this is
    // where Crashpad writes its local .dmp files (best-effort — without the
    // pin the platform default under userData is still local-only)
    deps.app.setPath('crashDumps', dumpsDir)
  } catch {
    // keep going: the platform default directory is still under userData
  }
  try {
    mkdirSync(dumpsDir, { recursive: true })
  } catch {
    // Crashpad recreates it when a crash lands
  }
  try {
    deps.crashReporter.start({
      // no submitURL + uploadToServer:false = crash reports never leave the
      // machine; they only land as .dmp files under dumpsDir
      uploadToServer: false,
      compress: true,
    })
  } catch {
    // a failed reporter start (exotic sandbox, unsupported platform) must
    // not stop the app; the ring log still captures the fatal events
  }
  return { dumpsDir }
}

/**
 * Log the fatal process events as structured single lines: every renderer
 * crash (the app-level event fires for the Home renderer AND every editor
 * tab's view), every gone child process (GPU/utility/network), and the
 * graceful quit boundary — a marker distinguishing "the user quit" from a
 * process that simply vanished without reaching will-quit.
 */
export function installFatalEventLogging(deps: {
  app: Pick<AppEventTarget, 'on'>
  log: EventLog
}): void {
  deps.app.on('render-process-gone', (...args: unknown[]) => {
    const webContents = args[1] as { id?: number } | undefined
    const details = args[2] as { reason?: string; exitCode?: number } | undefined
    deps.log.event('render-process-gone', {
      wc: webContents?.id,
      reason: details?.reason,
      exitCode: details?.exitCode,
    })
  })
  deps.app.on('child-process-gone', (...args: unknown[]) => {
    const details = args[1] as
      { type?: string; name?: string; reason?: string; exitCode?: number } | undefined
    deps.log.event('child-process-gone', {
      type: details?.type,
      name: details?.name,
      reason: details?.reason,
      exitCode: details?.exitCode,
    })
  })
  deps.app.on('will-quit', () => {
    deps.log.event('app-quit', {})
  })
}

export interface ShellDiagnostics {
  log: MainLog
  dumpsDir: string
}

let initialized: ShellDiagnostics | null = null

/**
 * One-shot wiring used by the shell main entry, placed BEFORE the heavy
 * startup phases (runtime configuration, IPC, windows): ring log under
 * userData/logs, console tee (records all existing console.error
 * diagnostics that used to be stderr-only), the local-only crash reporter,
 * and the fatal-event loggers. The process-level unhandledRejection /
 * uncaughtException handlers in the entry already console.error their
 * reason — the tee captures both into the log file.
 */
export function initShellDiagnostics(options: { userDataDir: string }): ShellDiagnostics {
  if (initialized) return initialized
  const logsDir = join(options.userDataDir, LOGS_DIR_NAME)
  const log = createMainLog({ dir: logsDir })
  teeConsoleInto(log)
  const { dumpsDir } = startLocalCrashReporter({
    app,
    crashReporter,
    userDataDir: options.userDataDir,
  })
  installFatalEventLogging({ app, log })
  log.event('crash-diagnostics-start', {
    uploadToServer: false,
    dumpsDir,
    logsDir,
    electron: process.versions.electron,
    shell: app.getVersion(),
  })
  initialized = { log, dumpsDir }
  return initialized
}

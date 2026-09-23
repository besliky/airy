import { mkdirSync, appendFileSync, renameSync, statSync, unlinkSync } from 'node:fs'
import { join } from 'node:path'
import { format as formatValues } from 'node:util'

/**
 * Crash-diagnostics ring log for the main process (UX-1628): a bounded
 * in-memory ring of recent lines plus a size-capped on-disk copy under
 * userData/logs, so main-process diagnostics survive a desktop-launch
 * (stderr from a .desktop file is lost) and remain available for a user
 * to attach to a "the app closed itself" report. Local files only — the
 * log never leaves the machine.
 *
 * Failure containment contract: the logger is on the crash path of the app
 * (fatal events are logged), so NO operation here may ever throw. Every
 * filesystem call is guarded; write failures are counted and silently
 * dropped, and the in-memory ring keeps working when the disk does not.
 */

export type MainLogLevel = 'debug' | 'info' | 'warn' | 'error'

export interface MainLogOptions {
  /** directory holding the log files; created recursively on init */
  dir: string
  /** active file name; rotated backups get `.1`, `.2` … suffixes */
  fileName?: string
  /** rotate when the active file would exceed this size, in bytes */
  maxFileBytes?: number
  /** rotated backups kept on disk besides the active file */
  maxBackups?: number
  /** in-memory ring bound (lines) */
  ringCapacity?: number
  /** persisted lines longer than this are truncated (paths, stack noise) */
  maxLineLength?: number
  /** injectable clock (tests) */
  now?: () => Date
}

export interface MainLog {
  /** append one pre-formatted line at a level */
  line(level: MainLogLevel, message: string): void
  /** structured single-line record for a named fatal/lifecycle event */
  event(name: string, fields?: Record<string, unknown>): void
  debug(...values: unknown[]): void
  info(...values: unknown[]): void
  warn(...values: unknown[]): void
  error(...values: unknown[]): void
  /** the newest ring lines, oldest first */
  recentLines(): string[]
  /** absolute paths of the log directory and the active file */
  paths(): { dir: string; activeFile: string }
  /** observability for tests and diagnostics: counters and rotation state */
  stats(): {
    ringSize: number
    activeFileBytes: number
    rotations: number
    droppedWrites: number
  }
}

export const MAIN_LOG_DEFAULTS = {
  fileName: 'main.log',
  /** one file per MiB keeps each dump small enough to eyeball and attach */
  maxFileBytes: 1024 * 1024,
  maxBackups: 3,
  ringCapacity: 2000,
  maxLineLength: 4000,
} as const

export function createMainLog(options: MainLogOptions): MainLog {
  const fileName = options.fileName ?? MAIN_LOG_DEFAULTS.fileName
  const maxFileBytes = Math.max(1, options.maxFileBytes ?? MAIN_LOG_DEFAULTS.maxFileBytes)
  const maxBackups = Math.max(0, options.maxBackups ?? MAIN_LOG_DEFAULTS.maxBackups)
  const ringCapacity = Math.max(1, options.ringCapacity ?? MAIN_LOG_DEFAULTS.ringCapacity)
  const maxLineLength = Math.max(32, options.maxLineLength ?? MAIN_LOG_DEFAULTS.maxLineLength)
  const now = options.now ?? (() => new Date())

  const dir = options.dir
  const activeFile = join(dir, fileName)

  // the in-memory ring: the last `ringCapacity` lines, oldest first
  const ring: string[] = []
  let activeSize = 0
  let rotations = 0
  let droppedWrites = 0

  // Guarded init: an unusable log directory (permissions, read-only volume)
  // must not stop the app from starting — writes then just get dropped.
  try {
    mkdirSync(dir, { recursive: true })
    activeSize = statSync(activeFile).size
  } catch {
    // missing active file is the normal first-run case; statSync on it lands here
    activeSize = 0
  }
  if (activeSize > maxFileBytes) rotate()

  function pushRing(line: string): void {
    ring.push(line)
    if (ring.length > ringCapacity) ring.splice(0, ring.length - ringCapacity)
  }

  function rotate(): void {
    try {
      // drop the oldest backup, shift the others up, retire the active file
      unlinkSync(join(dir, `${fileName}.${maxBackups}`))
    } catch {
      // no such backup — normal for young logs
    }
    for (let i = maxBackups - 1; i >= 1; i--) {
      try {
        renameSync(join(dir, `${fileName}.${i}`), join(dir, `${fileName}.${i + 1}`))
      } catch {
        // absent slot — nothing to shift
      }
    }
    try {
      renameSync(activeFile, join(dir, `${fileName}.1`))
      activeSize = 0
      rotations++
    } catch {
      // an unretirable active file (locked by AV scan, exotic fs) — keep
      // appending to it; the cap retriggers on the next oversized write
    }
  }

  function append(line: string): void {
    // rotation check against the would-be size, not the current one, so a
    // single oversized line can never balloon the active file past the cap
    if (activeSize + line.length + 1 > maxFileBytes) rotate()
    try {
      appendFileSync(activeFile, `${line}\n`)
      activeSize += line.length + 1
    } catch {
      droppedWrites++
    }
  }

  function line(level: MainLogLevel, message: string): void {
    const timestamp = now().toISOString()
    const flat = message.replace(/\r?\n/g, ' ')
    const text =
      flat.length > maxLineLength ? `${flat.slice(0, maxLineLength)}…(+${flat.length})` : flat
    const lineText = `${timestamp} ${level.toUpperCase()} ${text}`
    pushRing(lineText)
    append(lineText)
  }

  function render(values: unknown[]): string {
    try {
      return formatValues(...values)
    } catch {
      // a throwing custom inspector on some argument must not kill the log
      return values
        .map((v) => {
          try {
            return String(v)
          } catch {
            return '[unprintable]'
          }
        })
        .join(' ')
    }
  }

  function event(name: string, fields?: Record<string, unknown>): void {
    const parts = Object.entries(fields ?? {}).map(([key, value]) => {
      let text: string
      try {
        // JSON.stringify returns undefined for functions/symbols — fall back
        // to String(); a circular structure lands in the catch below
        const json = value === undefined ? undefined : JSON.stringify(value)
        text = json === undefined ? String(value) : json
      } catch {
        text = '[unserializable]'
      }
      return `${key}=${text}`
    })
    line('warn', `event ${name}${parts.length > 0 ? ` ${parts.join(' ')}` : ''}`)
  }

  return {
    line,
    event,
    debug: (...values) => line('debug', render(values)),
    info: (...values) => line('info', render(values)),
    warn: (...values) => line('warn', render(values)),
    error: (...values) => line('error', render(values)),
    recentLines: () => [...ring],
    paths: () => ({ dir, activeFile }),
    stats: () => ({
      ringSize: ring.length,
      activeFileBytes: activeSize,
      rotations,
      droppedWrites,
    }),
  }
}

/**
 * Mirror everything the main process already prints to stderr (console.*)
 * into the ring log. The shell has years of console.error diagnostics on
 * crash-adjacent paths; teeing them captures all of it without touching
 * those call sites. Returns a restore function (tests).
 *
 * The tee re-enters guarded: a console call made while a line is being
 * logged is dropped instead of recursing.
 */
export function teeConsoleInto(log: MainLog): () => void {
  const original = {
    debug: console.debug.bind(console),
    log: console.log.bind(console),
    info: console.info.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
  }
  let inside = false
  const tee =
    (level: MainLogLevel, write: (...values: unknown[]) => void) =>
    (...values: unknown[]): void => {
      // stderr keeps its original behavior first — the tee must not change
      // what a developer sees in a terminal
      write(...values)
      if (inside) return
      inside = true
      try {
        log.line(level, formatValues(...values))
      } catch {
        // never let the mirror break the real call
      } finally {
        inside = false
      }
    }
  console.debug = tee('debug', original.debug)
  console.log = tee('info', original.log)
  console.info = tee('info', original.info)
  console.warn = tee('warn', original.warn)
  console.error = tee('error', original.error)
  return () => {
    console.debug = original.debug
    console.log = original.log
    console.info = original.info
    console.warn = original.warn
    console.error = original.error
  }
}

import { randomUUID } from 'node:crypto'
import { readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import type { Rectangle } from 'electron'

/**
 * Persisted shell-window geometry (userData/window-state.json): the normal
 * (pre-maximize) bounds plus the maximize/fullscreen flags, saved on close and
 * (debounced) on move/resize so a relaunch reopens where the user left it.
 * The parse/validate logic is pure and unit-tested; the shell applies a state
 * only when it is fully on-screen (see {@link isWindowOnScreen}).
 */
export interface WindowState {
  x: number
  y: number
  width: number
  height: number
  isMaximized: boolean
  isFullScreen: boolean
}

/** sane floors mirroring the shell window's minWidth/minHeight */
const MIN_WIDTH = 720
const MIN_HEIGHT = 550
/** upper bound for a plausible display area today (8K plus overshoot) */
const MAX_DIMENSION = 16384

/**
 * Parse a raw JSON value (already JSON.parse-d) into a WindowState. Anything
 * malformed — missing fields, non-finite numbers, sub-minimum or absurd
 * sizes — yields null so the caller falls back to the default geometry.
 */
export function parseWindowState(raw: unknown): WindowState | null {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null
  const record = raw as Record<string, unknown>
  const bounds = record.bounds
  if (bounds === null || typeof bounds !== 'object' || Array.isArray(bounds)) return null
  const b = bounds as Record<string, unknown>
  const { x, y, width, height } = b
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    typeof x !== 'number' ||
    typeof y !== 'number' ||
    typeof width !== 'number' ||
    typeof height !== 'number'
  ) {
    return null
  }
  if (width < MIN_WIDTH || height < MIN_HEIGHT) return null
  if (width > MAX_DIMENSION || height > MAX_DIMENSION) return null
  if (Math.abs(x) > MAX_DIMENSION || Math.abs(y) > MAX_DIMENSION) return null
  return {
    x: Math.round(x),
    y: Math.round(y),
    width: Math.round(width),
    height: Math.round(height),
    isMaximized: record.isMaximized === true,
    isFullScreen: record.isFullScreen === true,
  }
}

/**
 * A restored window must land fully inside a single current display — a state
 * saved on a monitor that has since been unplugged (or corrupted coordinates)
 * is discarded rather than reopening an unreachable window. A window spanning
 * two displays is likewise not restored (it belongs to no one display); the
 * default geometry opens instead and the next close persists fresh bounds.
 */
export function isWindowOnScreen(state: WindowState, displays: Rectangle[]): boolean {
  const window: Rectangle = { x: state.x, y: state.y, width: state.width, height: state.height }
  return displays.some(
    (display) =>
      window.x >= display.x &&
      window.y >= display.y &&
      window.x + window.width <= display.x + display.width &&
      window.y + window.height <= display.y + display.height,
  )
}

/** Read and parse the state file; missing/corrupt content yields null. */
export function readWindowState(path: string): WindowState | null {
  let raw: unknown
  try {
    raw = JSON.parse(readFileSync(path, 'utf8'))
  } catch {
    return null // never shipped / not JSON — default geometry
  }
  return parseWindowState(raw)
}

/** Persist atomically (temp + rename, same pattern as app-settings.json). */
export function writeWindowState(path: string, state: WindowState): void {
  const payload = {
    bounds: { x: state.x, y: state.y, width: state.width, height: state.height },
    isMaximized: state.isMaximized,
    isFullScreen: state.isFullScreen,
  }
  const tempPath = `${path}.${process.pid}.${randomUUID()}.tmp`
  try {
    writeFileSync(tempPath, JSON.stringify(payload, null, 2), {
      encoding: 'utf8',
      flag: 'wx',
      flush: true,
    })
    renameSync(tempPath, path)
  } catch (error) {
    try {
      unlinkSync(tempPath)
    } catch {
      // The write may have failed before the temporary file was created.
    }
    throw error
  }
}

/**
 * Move/resize events arrive in bursts, so the write they trigger is
 * debounced — but short (UX-1775): a kill -9 inside the debounce window
 * loses the last move, and half a second is far too much window to lose.
 * At this width the file trails the window by at most ~150 ms, and the
 * close/maximize paths below still flush synchronously.
 */
export const GEOMETRY_SAVE_DEBOUNCE_MS = 150

/** The three geometry-write entry points the primary window wires up. */
export interface GeometrySaver {
  /** coalesce a burst of move/resize events into one deferred write */
  schedule: () => void
  /** write NOW (window close, maximize transitions) and drop the pending one */
  flush: () => void
  /** drop a pending deferred write without writing */
  cancel: () => void
}

/**
 * Debounced geometry saver around a persist callback (pure, unit-tested):
 * a burst of `schedule()` calls collapses into a single write
 * {@link GEOMETRY_SAVE_DEBOUNCE_MS} after the last event; `flush()` writes
 * immediately and cancels the pending deferred write so nothing lands after
 * the synchronous one.
 */
export function createGeometrySaver(
  persist: () => void,
  debounceMs: number = GEOMETRY_SAVE_DEBOUNCE_MS,
): GeometrySaver {
  let timer: ReturnType<typeof setTimeout> | null = null
  const cancel = (): void => {
    if (timer !== null) {
      clearTimeout(timer)
      timer = null
    }
  }
  return {
    schedule: () => {
      cancel()
      timer = setTimeout(() => {
        timer = null
        persist()
      }, debounceMs)
    },
    flush: () => {
      cancel()
      persist()
    },
    cancel,
  }
}

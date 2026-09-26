import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Window-state persistence (src/main/window-state.ts): the pure parse and
 * on-screen validation logic that guards restoring the shell window's
 * geometry, the atomic file round-trip, and the debounced geometry saver
 * (UX-1775: a kill -9 inside the debounce window must not lose the move).
 */

let WindowState: typeof import('../src/main/window-state')
let scratch: string

beforeEach(async () => {
  WindowState = await import('../src/main/window-state')
  scratch = mkdtempSync(join(tmpdir(), 'airy-window-state-'))
})

afterEach(() => {
  rmSync(scratch, { recursive: true, force: true })
})

const DISPLAY = { x: 0, y: 0, width: 1920, height: 1080 }
const DISPLAY_2 = { x: 1920, y: 0, width: 2560, height: 1440 }

describe('parseWindowState', () => {
  it('parses a well-formed saved state', () => {
    const state = WindowState.parseWindowState({
      bounds: { x: 12, y: 34.6, width: 1360.4, height: 900 },
      isMaximized: true,
      isFullScreen: false,
    })
    // coordinates round to integers — Electron bounds are whole pixels
    expect(state).toEqual({
      x: 12,
      y: 35,
      width: 1360,
      height: 900,
      isMaximized: true,
      isFullScreen: false,
    })
  })

  it('defaults both flags to false when absent', () => {
    const state = WindowState.parseWindowState({
      bounds: { x: 0, y: 0, width: 800, height: 600 },
    })
    expect(state?.isMaximized).toBe(false)
    expect(state?.isFullScreen).toBe(false)
  })

  it('rejects malformed payloads', () => {
    for (const bad of [
      null,
      undefined,
      'garbage',
      42,
      [],
      {},
      { bounds: null },
      { bounds: 'x' },
      { bounds: { x: 0, y: 0 } },
      { bounds: { x: '0', y: 0, width: 800, height: 600 } },
      { bounds: { x: Number.NaN, y: 0, width: 800, height: 600 } },
      { bounds: { x: Infinity, y: 0, width: 800, height: 600 } },
    ]) {
      expect(WindowState.parseWindowState(bad)).toBeNull()
    }
  })

  it('rejects implausible sizes', () => {
    // below the window's minimum chrome size
    expect(
      WindowState.parseWindowState({ bounds: { x: 0, y: 0, width: 100, height: 100 } }),
    ).toBeNull()
    // beyond any current display (corrupt / hostile values)
    expect(
      WindowState.parseWindowState({ bounds: { x: 0, y: 0, width: 999999, height: 900 } }),
    ).toBeNull()
    expect(
      WindowState.parseWindowState({ bounds: { x: 999999, y: 0, width: 900, height: 900 } }),
    ).toBeNull()
  })
})

describe('isWindowOnScreen', () => {
  it('accepts a window fully inside one display', () => {
    const state = {
      x: 100,
      y: 100,
      width: 1360,
      height: 900,
      isMaximized: false,
      isFullScreen: false,
    }
    expect(WindowState.isWindowOnScreen(state, [DISPLAY, DISPLAY_2])).toBe(true)
  })

  it('accepts a window on the second display', () => {
    const state = {
      x: 2000,
      y: 50,
      width: 1360,
      height: 900,
      isMaximized: false,
      isFullScreen: false,
    }
    expect(WindowState.isWindowOnScreen(state, [DISPLAY, DISPLAY_2])).toBe(true)
  })

  it('rejects a window on a disconnected display', () => {
    const state = {
      x: 2000,
      y: 50,
      width: 1360,
      height: 900,
      isMaximized: false,
      isFullScreen: false,
    }
    expect(WindowState.isWindowOnScreen(state, [DISPLAY])).toBe(false)
  })

  it('rejects a window hanging off the display edge', () => {
    // only half of the window is reachable — discard, do not restore clipped
    const state = {
      x: 1400,
      y: 0,
      width: 1360,
      height: 900,
      isMaximized: false,
      isFullScreen: false,
    }
    expect(WindowState.isWindowOnScreen(state, [DISPLAY])).toBe(false)
  })

  it('rejects a window straddling two displays (belongs to neither)', () => {
    const state = {
      x: 1500,
      y: 0,
      width: 1360,
      height: 900,
      isMaximized: false,
      isFullScreen: false,
    }
    expect(WindowState.isWindowOnScreen(state, [DISPLAY, DISPLAY_2])).toBe(false)
  })

  it('rejects with no displays at all', () => {
    const state = { x: 0, y: 0, width: 800, height: 600, isMaximized: false, isFullScreen: false }
    expect(WindowState.isWindowOnScreen(state, [])).toBe(false)
  })
})

describe('state file round-trip', () => {
  it('writes atomically and reads back the same state', () => {
    const path = join(scratch, 'window-state.json')
    WindowState.writeWindowState(path, {
      x: 10,
      y: 20,
      width: 1360,
      height: 900,
      isMaximized: true,
      isFullScreen: false,
    })
    // the persisted shape is the nested bounds object
    expect(JSON.parse(readFileSync(path, 'utf8'))).toEqual({
      bounds: { x: 10, y: 20, width: 1360, height: 900 },
      isMaximized: true,
      isFullScreen: false,
    })
    expect(WindowState.readWindowState(path)).toEqual({
      x: 10,
      y: 20,
      width: 1360,
      height: 900,
      isMaximized: true,
      isFullScreen: false,
    })
  })

  it('reads a missing or corrupt file as null (default geometry)', () => {
    expect(WindowState.readWindowState(join(scratch, 'absent.json'))).toBeNull()
    const corrupt = join(scratch, 'window-state.json')
    writeFileSync(corrupt, '{ not json', 'utf8')
    expect(WindowState.readWindowState(corrupt)).toBeNull()
  })

  it('leaves the previous file intact when a write fails mid-way', () => {
    const path = join(scratch, 'nested', 'window-state.json') // parent dir absent
    WindowState.writeWindowState(join(scratch, 'window-state.json'), {
      x: 0,
      y: 0,
      width: 800,
      height: 600,
      isMaximized: false,
      isFullScreen: false,
    })
    expect(() =>
      WindowState.writeWindowState(path, {
        x: 1,
        y: 1,
        width: 800,
        height: 600,
        isMaximized: false,
        isFullScreen: false,
      }),
    ).toThrow()
    expect(WindowState.readWindowState(join(scratch, 'window-state.json'))?.x).toBe(0)
  })
})

// ── UX-1775: the debounced geometry saver ──
// The audit (SET-26-5): with a 500 ms debounce, kill -9 right after a move
// lost it. The debounce is now short, and close/maximize flush synchronously.

describe('GEOMETRY_SAVE_DEBOUNCE_MS', () => {
  it('is short enough that a kill in the window loses almost nothing', () => {
    expect(WindowState.GEOMETRY_SAVE_DEBOUNCE_MS).toBe(150)
  })
})

describe('createGeometrySaver', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  it('coalesces a burst of moves into one deferred write', () => {
    const persist = vi.fn()
    const saver = WindowState.createGeometrySaver(persist)
    for (let i = 0; i < 10; i++) {
      saver.schedule()
      vi.advanceTimersByTime(50)
    }
    expect(persist).not.toHaveBeenCalled()
    vi.advanceTimersByTime(WindowState.GEOMETRY_SAVE_DEBOUNCE_MS)
    expect(persist).toHaveBeenCalledTimes(1)
  })

  it('writes the final state synchronously on flush and nothing lands after', () => {
    const persist = vi.fn()
    const saver = WindowState.createGeometrySaver(persist)
    saver.schedule()
    saver.flush() // the close path: one immediate write …
    expect(persist).toHaveBeenCalledTimes(1)
    vi.advanceTimersByTime(10_000) // … and no deferred write afterwards
    expect(persist).toHaveBeenCalledTimes(1)
  })

  it('flush without a pending write still persists once (close persists)', () => {
    const persist = vi.fn()
    const saver = WindowState.createGeometrySaver(persist)
    saver.flush()
    expect(persist).toHaveBeenCalledTimes(1)
  })

  it('cancel drops a pending deferred write without persisting', () => {
    const persist = vi.fn()
    const saver = WindowState.createGeometrySaver(persist)
    saver.schedule()
    saver.cancel()
    vi.advanceTimersByTime(10_000)
    expect(persist).not.toHaveBeenCalled()
  })

  it('schedules again after a flush (move after maximize)', () => {
    const persist = vi.fn()
    const saver = WindowState.createGeometrySaver(persist)
    saver.flush()
    saver.schedule()
    vi.advanceTimersByTime(WindowState.GEOMETRY_SAVE_DEBOUNCE_MS)
    expect(persist).toHaveBeenCalledTimes(2)
  })

  it('honors a custom debounce (injectable for tests)', () => {
    const persist = vi.fn()
    const saver = WindowState.createGeometrySaver(persist, 1000)
    saver.schedule()
    vi.advanceTimersByTime(999)
    expect(persist).not.toHaveBeenCalled()
    vi.advanceTimersByTime(1)
    expect(persist).toHaveBeenCalledTimes(1)
  })
})

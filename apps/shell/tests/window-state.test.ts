import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

/**
 * Window-state persistence (src/main/window-state.ts): the pure parse and
 * on-screen validation logic that guards restoring the shell window's
 * geometry, plus the atomic file round-trip.
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

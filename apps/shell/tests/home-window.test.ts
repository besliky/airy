// PERF-1736: windowing math for the Home file tables. Pure arithmetic, so the
// clamping, overscan and unmeasured-viewport fallback rules are all tested
// here without a DOM.
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_ROW_PITCH,
  MIN_WINDOW_ROWS,
  OVERSCAN_ROWS,
  spacerHeights,
  VIRTUALIZE_THRESHOLD,
  windowFor,
} from '../src/renderer/src/home-window'

describe('windowFor', () => {
  it('renders the viewport plus overscan on both sides', () => {
    // 10 rows visible around index 100 → start 100-8, window = max(10+16, 48)
    const win = windowFor({ firstVisible: 100.4, visibleRows: 10, rowCount: 20_000 })
    expect(win.start).toBe(100 - OVERSCAN_ROWS)
    expect(win.end).toBe(win.start + MIN_WINDOW_ROWS)
  })

  it('grows the window when the viewport shows more than the minimum', () => {
    const win = windowFor({ firstVisible: 500, visibleRows: 80, rowCount: 20_000 })
    expect(win.start).toBe(500 - OVERSCAN_ROWS)
    expect(win.end).toBe(win.start + 80 + 2 * OVERSCAN_ROWS)
  })

  it('clamps to the list bounds on both edges', () => {
    const top = windowFor({ firstVisible: 0, visibleRows: 10, rowCount: 100 })
    expect(top.start).toBe(0)
    const bottom = windowFor({ firstVisible: 99, visibleRows: 10, rowCount: 100 })
    expect(bottom.end).toBe(100)
    expect(bottom.start).toBeLessThan(100)
  })

  it('never asks for rows past the end of the list', () => {
    const win = windowFor({ firstVisible: 19_990, visibleRows: 10, rowCount: 20_000 })
    expect(win.end).toBe(20_000)
    expect(win.start).toBeLessThan(20_000)
  })

  it('covers a tiny list completely', () => {
    const win = windowFor({ firstVisible: 3, visibleRows: 10, rowCount: 12 })
    expect(win).toEqual({ start: 0, end: 12 })
  })

  it('falls back to the minimum window when nothing is measurable', () => {
    const win = windowFor({ firstVisible: 0, visibleRows: 0, rowCount: 20_000 })
    expect(win).toEqual({ start: 0, end: MIN_WINDOW_ROWS })
  })

  it('tolerates negative firstVisible (list top below the viewport)', () => {
    const win = windowFor({ firstVisible: -12, visibleRows: 10, rowCount: 20_000 })
    expect(win.start).toBe(0)
  })

  it('handles an empty list', () => {
    expect(windowFor({ firstVisible: 0, visibleRows: 0, rowCount: 0 })).toEqual({
      start: 0,
      end: 0,
    })
  })
})

describe('spacerHeights', () => {
  it('splits the unrendered space above and below the window', () => {
    const pitch = DEFAULT_ROW_PITCH
    const win = windowFor({ firstVisible: 100, visibleRows: 10, rowCount: 1_000 })
    const { top, bottom } = spacerHeights(win, 1_000, pitch)
    expect(top).toBe(win.start * pitch)
    expect(bottom).toBe((1_000 - win.end) * pitch)
    // spacers + window reconstruct the full list height exactly
    expect(top + bottom + (win.end - win.start) * pitch).toBe(1_000 * pitch)
  })

  it('is zero on both sides for a fully rendered list', () => {
    const win = windowFor({ firstVisible: 0, visibleRows: 10, rowCount: 12 })
    expect(spacerHeights(win, 12, 53)).toEqual({ top: 0, bottom: 0 })
  })
})

describe('constants', () => {
  it('keep the virtualized window far below the catalog cap', () => {
    // the whole point of PERF-1736: 20k rows must never all render
    expect(VIRTUALIZE_THRESHOLD).toBeLessThan(1_000)
    expect(MIN_WINDOW_ROWS + 2 * OVERSCAN_ROWS).toBeLessThan(100)
    expect(DEFAULT_ROW_PITCH).toBeGreaterThan(0)
  })
})

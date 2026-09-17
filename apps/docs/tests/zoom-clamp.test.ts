/** The 50–400 zoom window every docs zoom input funnels through
 *  (src/renderer/zoom-clamp.ts): wheel deltas, the page-fit floor, and the
 *  ±10 step buttons share one clamp so they can never disagree. */
import { describe, expect, it } from 'vitest'

import { clampDocZoom, MAX_DOC_ZOOM, MIN_DOC_ZOOM } from '../src/renderer/zoom-clamp'

describe('clampDocZoom', () => {
  it('keeps in-window values untouched, including fractional wheel deltas', () => {
    expect(clampDocZoom(100)).toBe(100)
    expect(clampDocZoom(50)).toBe(50)
    expect(clampDocZoom(400)).toBe(400)
    expect(clampDocZoom(83.4)).toBe(83.4)
  })

  it('clamps overshoots back to the window bounds', () => {
    expect(clampDocZoom(10)).toBe(MIN_DOC_ZOOM)
    expect(clampDocZoom(-5)).toBe(MIN_DOC_ZOOM)
    expect(clampDocZoom(401)).toBe(MAX_DOC_ZOOM)
    expect(clampDocZoom(10000)).toBe(MAX_DOC_ZOOM)
  })

  it('falls back to 100 for non-finite input (fit computation failure)', () => {
    expect(clampDocZoom(NaN)).toBe(100)
    expect(clampDocZoom(Number.POSITIVE_INFINITY)).toBe(100)
  })
})

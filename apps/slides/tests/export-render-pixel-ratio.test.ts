import { describe, expect, it, vi } from 'vitest'

// The fake stage captures what the renderer passes to toBlob: a per-slide
// pixelRatio array must reach each slide's own encode (BUG-1302), not the
// first slide's scalar.
const toBlobCalls: Array<{ pixelRatio: number }> = []

// SlideThumb is stubbed so the offscreen mount resolves immediately with a
// stage that records its toBlob options (no Konva/canvas needed)
vi.mock('../src/renderer/SlideThumb', () => ({
  SlideThumb: (props: { slide: { id: number }; stageRef?: (stage: unknown) => void }) => {
    props.stageRef?.({
      toBlob: async (opts: { pixelRatio: number }) => {
        toBlobCalls.push(opts)
        return new Blob(['png'])
      },
    })
    return null
  },
}))

import type { RenderSlide } from '@airy-office/pptx-render'
import { createSlidePngRenderer } from '../src/renderer/export-render'

/**
 * BUG-1302: PR #84's render-cancel rework regressed BUG-1211 back to a single
 * scalar pixel ratio (the first slide's), so slides smaller than the first in
 * a mixed-size deck rendered under-sampled and drawImage-stretched (blur).
 * The renderer must again accept per-slide ratios and pass each slide its own.
 */
const slide = (i: number, widthPx: number) => ({ id: i, widthPx }) as unknown as RenderSlide

/** fresh capture list per test (the fake stage appends to the module-level array) */
const freshCalls = () => {
  toBlobCalls.length = 0
  return toBlobCalls
}

describe('createSlidePngRenderer per-slide pixel ratio (BUG-1211/1302)', () => {
  it('renders each slide at its own ratio when given an array', async () => {
    const calls = freshCalls()
    const renderer = createSlidePngRenderer(
      [slide(0, 960), slide(1, 480), slide(2, 960)],
      new Map(),
      [1, 0.5, 2],
    )
    try {
      await renderer.renderPng(0)
      await renderer.renderPng(1)
      await renderer.renderPng(2)
    } finally {
      renderer.dispose()
    }
    expect(calls.map((c) => c.pixelRatio)).toEqual([1, 0.5, 2])
  })

  it('falls back to the first ratio for an index the array does not cover', async () => {
    const calls = freshCalls()
    const renderer = createSlidePngRenderer([slide(0, 960), slide(1, 960)], new Map(), [2])
    try {
      await renderer.renderPng(0)
      await renderer.renderPng(1)
    } finally {
      renderer.dispose()
    }
    expect(calls.map((c) => c.pixelRatio)).toEqual([2, 2])
  })

  it('keeps the plain scalar path working', async () => {
    const calls = freshCalls()
    const renderer = createSlidePngRenderer([slide(0, 960), slide(1, 960)], new Map(), 3)
    try {
      await renderer.renderPng(0)
      await renderer.renderPng(1)
    } finally {
      renderer.dispose()
    }
    expect(calls.map((c) => c.pixelRatio)).toEqual([3, 3])
  })
})

import { describe, expect, it, vi } from 'vitest'

// react-konva's node entry requires the native 'canvas' package; the
// cancelled loop returns before any stage is created, nothing draws
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import type { RenderSlide } from '@airy-office/pptx-render'
import { renderSlidesToPngBase64 } from '../src/renderer/export-render'

/**
 * UX-1202: the video export's render phase honors the cooperative cancel
 * box between slides (the record loop already did). The pre-cancelled case
 * proves the gate runs before any Konva work — the loop returns an empty
 * result without rendering a single slide, and no progress leaks out.
 */
const slide = (i: number) => ({ id: i, widthPx: 960, heightPx: 540 }) as unknown as RenderSlide

describe('renderSlidesToPngBase64 cancel (UX-1202)', () => {
  it('stops before the first slide when already cancelled', async () => {
    const onProgress = vi.fn()
    const out = await renderSlidesToPngBase64(
      [slide(0), slide(1), slide(2)],
      new Map(),
      1,
      onProgress,
      { current: true },
    )
    expect(out).toEqual([]) // nothing rendered, nothing to record or write
    expect(onProgress).not.toHaveBeenCalled()
  })
})

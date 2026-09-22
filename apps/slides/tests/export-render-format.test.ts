import { describe, expect, it, vi } from 'vitest'

/**
 * PAR-316 JPEG export: renderSlidesToPngBase64 gained an output format —
 * 'jpeg' encodes through stage.toDataURL with the image/jpeg mime type (plus
 * a quality knob) while the default stays PNG; the base64 prefix is stripped
 * for both. The offscreen Konva Stage is faked here (react-konva needs the
 * native 'canvas' package), so only the export plumbing is exercised.
 */

const stageCalls: Array<{ mimeType: string; pixelRatio: number; quality?: number }> = []

vi.mock('react-konva', () => {
  const stub = () => null
  // React 19 passes `ref` in props for function components — SlideThumb's
  // ref callback receives the fake stage and resolves the draw promise.
  const Stage = (props: { ref?: (stage: unknown) => void; width?: number; height?: number }) => {
    queueMicrotask(() =>
      props.ref?.({
        toDataURL: (opts: { mimeType: string; pixelRatio: number; quality?: number }) => {
          stageCalls.push(opts)
          return `data:${opts.mimeType};base64,QUJD`
        },
      }),
    )
    return null
  }
  return {
    Stage,
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

// SlideThumb paints its page rect through fillToKonva, so the stub slide needs
// a real (solid) background fill — but no nodes, nothing else draws.
const slide = (i: number) =>
  ({
    id: i,
    widthPx: 960,
    heightPx: 540,
    background: { kind: 'solid', color: '#ffffff' },
    nodes: [],
  }) as unknown as RenderSlide

describe('renderSlidesToPngBase64 output format (PAR-316)', () => {
  it('defaults to PNG without a quality override', async () => {
    stageCalls.length = 0
    const out = await renderSlidesToPngBase64([slide(0)], new Map())
    expect(out).toEqual(['QUJD'])
    expect(stageCalls[0]!.mimeType).toBe('image/png')
    expect(stageCalls[0]!.quality).toBeUndefined()
  })

  it('encodes JPEG with quality 0.92 when requested', async () => {
    stageCalls.length = 0
    const out = await renderSlidesToPngBase64(
      [slide(0), slide(1)],
      new Map(),
      2,
      undefined,
      undefined,
      'jpeg',
    )
    expect(out).toEqual(['QUJD', 'QUJD'])
    for (const call of stageCalls) {
      expect(call.mimeType).toBe('image/jpeg')
      expect(call.quality).toBeCloseTo(0.92)
    }
  })
})

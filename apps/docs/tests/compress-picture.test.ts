import { describe, expect, it } from 'vitest'
import {
  compressedMime,
  compressPictureDataUrl,
  compressSourceWindow,
  compressTargetSize,
  type CompressEncoder,
  type SizeProbe,
} from '../src/renderer/editor/compress-picture'

/** fake encoder capturing the request, returning a shorter dataUrl */
const fakeEncoder =
  (factor: number): CompressEncoder =>
  async (dataUrl, _window, _target, mime) => {
    const mimeTag = mime === 'image/jpeg' ? 'data:image/jpeg' : 'data:image/png'
    return `${mimeTag};base64,${'A'.repeat(Math.max(1, Math.round(dataUrl.length * factor) - 32))}`
  }

/** stub probe: jsdom cannot decode images, tests hand the natural size over */
const probe =
  (size: { widthPx: number; heightPx: number }): SizeProbe =>
  async () =>
    size

describe('compressTargetSize', () => {
  it('resamples the display size to the chosen ppi', () => {
    // displayed 4x2 inch (384x192 px at 96dpi), natural 3840x1920
    const natural = { widthPx: 3840, heightPx: 1920 }
    const display = { widthPx: 384, heightPx: 192 }
    expect(compressTargetSize(natural, display, { ppi: 96, deleteCropped: false })).toEqual({
      widthPx: 384,
      heightPx: 192,
    })
    expect(compressTargetSize(natural, display, { ppi: 150, deleteCropped: false })).toEqual({
      widthPx: 600,
      heightPx: 300,
    })
    expect(compressTargetSize(natural, display, { ppi: 220, deleteCropped: false })).toEqual({
      widthPx: 880,
      heightPx: 440,
    })
  })

  it('never upscales beyond the natural size', () => {
    const out = compressTargetSize(
      { widthPx: 300, heightPx: 150 },
      { widthPx: 480, heightPx: 240 },
      { ppi: 220, deleteCropped: false },
    )
    expect(out).toEqual({ widthPx: 300, heightPx: 150 })
  })

  it('bakes the crop window when deleteCropped is set, keeps the source otherwise', () => {
    const crop = { l: 0.25, t: 0, r: 0.25, b: 0 } // keep the middle half
    const window = compressSourceWindow(
      { widthPx: 1000, heightPx: 400 },
      { ppi: 96, deleteCropped: true, crop },
    )
    expect(window).toEqual({ x: 250, y: 0, widthPx: 500, heightPx: 400 })
    const kept = compressSourceWindow(
      { widthPx: 1000, heightPx: 400 },
      { ppi: 96, deleteCropped: false, crop },
    )
    expect(kept).toEqual({ x: 0, y: 0, widthPx: 1000, heightPx: 400 })
  })

  it('scales the whole source to the visible window when the crop survives', () => {
    const crop = { l: 0.25, t: 0, r: 0.25, b: 0 }
    // visible 500px window shown at 192 px, 96 ppi → scale 0.384 on the source
    const out = compressTargetSize(
      { widthPx: 1000, heightPx: 400 },
      { widthPx: 192, heightPx: 192 },
      { ppi: 96, deleteCropped: false, crop },
    )
    expect(out).toEqual({ widthPx: 384, heightPx: 154 })
  })
})

describe('compressedMime', () => {
  it('keeps png for transparency-bearing sources and jpeg for photos', () => {
    expect(compressedMime('image/png')).toBe('image/png')
    expect(compressedMime('image/gif')).toBe('image/png')
    expect(compressedMime('image/jpeg')).toBe('image/jpeg')
  })
})

describe('compressPictureDataUrl', () => {
  it('returns the encoded result with the right content type and smaller bytes', async () => {
    const src = `data:image/png;base64,${'B'.repeat(4000)}`
    const out = await compressPictureDataUrl(
      src,
      { widthPx: 384, heightPx: 192 },
      { ppi: 96, deleteCropped: false },
      fakeEncoder(0.5),
      probe({ widthPx: 3840, heightPx: 1920 }),
    )
    expect(out).not.toBeNull()
    expect(out!.dataUrl.startsWith('data:image/png;base64,')).toBe(true)
    expect(out!.dataUrl.length).toBeLessThan(src.length)
    expect(out!.mime).toBe('image/png')
    expect(out!.widthPx).toBe(384)
  })

  it('produces a jpeg content type for jpeg sources', async () => {
    const src = `data:image/jpeg;base64,${'B'.repeat(4000)}`
    const out = await compressPictureDataUrl(
      src,
      { widthPx: 96, heightPx: 96 },
      { ppi: 96, deleteCropped: false },
      fakeEncoder(0.5),
      probe({ widthPx: 960, heightPx: 960 }),
    )
    expect(out!.mime).toBe('image/jpeg')
    expect(out!.dataUrl.startsWith('data:image/jpeg;base64,')).toBe(true)
  })

  it('rejects a re-encode that did not shrink', async () => {
    const src = `data:image/png;base64,${'B'.repeat(4000)}`
    const bigger = await compressPictureDataUrl(
      src,
      { widthPx: 384, heightPx: 192 },
      { ppi: 96, deleteCropped: false },
      fakeEncoder(2),
      probe({ widthPx: 3840, heightPx: 1920 }),
    )
    expect(bigger).toBeNull()
  })

  it('passes the crop window and target size to the encoder', async () => {
    const src = `data:image/png;base64,${'B'.repeat(4000)}`
    let seen: unknown = null
    const spy: CompressEncoder = async (_d, window, target, mime) => {
      seen = { window, target, mime }
      return `data:image/png;base64,${'A'.repeat(2000)}`
    }
    await compressPictureDataUrl(
      src,
      { widthPx: 192, heightPx: 120 },
      { ppi: 96, deleteCropped: true, crop: { l: 0.5, t: 0, r: 0, b: 0 } },
      spy,
      probe({ widthPx: 1920, heightPx: 1200 }),
    )
    // crop window: left half; target: the 960px window at 96 ppi of 192px display
    expect(seen).toEqual({
      window: { x: 960, y: 0, widthPx: 960, heightPx: 1200 },
      target: { widthPx: 192, heightPx: 240 },
      mime: 'image/png',
    })
  })

  it('rejects unsupported source types', async () => {
    const out = await compressPictureDataUrl(
      'data:image/webp;base64,AAAA',
      { widthPx: 100, heightPx: 100 },
      { ppi: 96, deleteCropped: false },
    )
    expect(out).toBeNull()
  })
})

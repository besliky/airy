/**
 * Offscreen rendering for export — draws each RenderSlide page into an offscreen Konva Stage
 * and exports high-resolution PNGs. Reuses SlideThumb (same rendering as the main canvas and
 * thumbnails), guaranteeing the export matches what the editor shows.
 */
import React from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type Konva from 'konva'
import type { RenderSlide } from '@airy-office/pptx-render'
import { SlideThumb } from './SlideThumb'

/** Pixel ratio of the exported bitmap (2x hi-res, 1280 viewport width → 2560px PNG) */
const EXPORT_PIXEL_RATIO = 2

/**
 * Draw `slide` into a stage on the shared offscreen root and wait for Konva's
 * batchDraw: resolves the stage once it holds the slide's content.
 */
function drawSlideStage(
  root: Root,
  slide: RenderSlide,
  images: Map<string, HTMLImageElement>,
): Promise<Konva.Stage> {
  return new Promise<Konva.Stage>((resolve) => {
    root.render(
      <SlideThumb
        slide={slide}
        images={images}
        width={slide.widthPx}
        stageRef={(s) => s && resolve(s)}
      />,
    )
  })
}

/** Wait one frame so Konva's batchDraw finishes before the bitmap is captured. */
const nextFrame = () => new Promise((r) => requestAnimationFrame(r))

/**
 * Render each page to PNG base64 (without the data: prefix).
 * Reuse a single offscreen root page by page, grabbing each page as it's drawn.
 * pixelRatio 1 is enough for AI-vision screenshots (half the tokens of the 2x export default).
 * `cancel` (cooperative, checked between slides — the images/PDF export phase)
 * stops the loop early and returns the slides drawn so far.
 */
export async function renderSlidesToPngBase64(
  slides: RenderSlide[],
  images: Map<string, HTMLImageElement>,
  pixelRatio: number | ReadonlyArray<number> = EXPORT_PIXEL_RATIO,
  onProgress?: (done: number, total: number) => void,
  cancel?: { current: boolean },
): Promise<string[]> {
  // Offscreen container: mounted outside the body viewport (display:none would give the Konva canvas zero size, unusable)
  const container = document.createElement('div')
  container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(container)
  const root = createRoot(container)
  const out: string[] = []
  try {
    for (const slide of slides) {
      if (cancel?.current) break // UX-1202: Cancel must work during the render phase too
      const stage = await drawSlideStage(root, slide, images)
      await nextFrame()
      const ratio =
        typeof pixelRatio === 'number'
          ? pixelRatio
          : (pixelRatio[out.length] ?? pixelRatio[0] ?? EXPORT_PIXEL_RATIO)
      const dataUrl = stage.toDataURL({ mimeType: 'image/png', pixelRatio: ratio })
      out.push(dataUrl.replace(/^data:image\/png;base64,/, ''))
      onProgress?.(out.length, slides.length)
    }
  } finally {
    root.unmount()
    container.remove()
  }
  return out
}

/**
 * On-demand PNG renderer for the video export (BUG-1300): slides are rendered
 * one at a time, on request, as PNG *blobs* — never as base64 strings, and
 * never accumulated. The caller (the recording loop) asks for item i, gets
 * its blob, decodes it, and lets both go; the deck never sits in memory.
 */
export interface SlidePngRenderer {
  /** Render the slide at `index` (of the slides passed in) to a PNG blob. */
  renderPng(index: number): Promise<Blob>
  /** Tear down the offscreen root — always call from a finally. */
  dispose(): void
}

/**
 * Create the on-demand renderer: one offscreen container/root reused for
 * every slide (same reuse discipline as renderSlidesToPngBase64), mounted
 * eagerly and disposed by the caller. renderPng is sequential by contract —
 * the recording loop awaits each call (plus at most one in-flight prefetch).
 * `pixelRatio` may be per-slide (BUG-1211): in mixed-size decks each slide
 * renders at its own fit scale, so no PNG is over- or under-sampled for its
 * box (an index missing from the array falls back to the first entry).
 */
export function createSlidePngRenderer(
  slides: ReadonlyArray<RenderSlide>,
  images: Map<string, HTMLImageElement>,
  pixelRatio: number | ReadonlyArray<number> = EXPORT_PIXEL_RATIO,
): SlidePngRenderer {
  const container = document.createElement('div')
  container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(container)
  const root = createRoot(container)
  let disposed = false
  return {
    async renderPng(index: number) {
      const slide = slides[index]
      if (disposed || !slide) throw new Error('slide PNG renderer is disposed or out of range')
      const stage = await drawSlideStage(root, slide, images)
      await nextFrame()
      const ratio =
        typeof pixelRatio === 'number'
          ? pixelRatio
          : (pixelRatio[index] ?? pixelRatio[0] ?? EXPORT_PIXEL_RATIO)
      // toBlob lands the compressed PNG in blob storage — no base64 string is
      // ever built, so nothing deck-sized enters the JS heap
      const blob = (await stage.toBlob({ mimeType: 'image/png', pixelRatio: ratio })) as Blob | null
      if (!blob) throw new Error('slide PNG encode failed')
      return blob
    },
    dispose() {
      if (disposed) return
      disposed = true
      root.unmount()
      container.remove()
    },
  }
}

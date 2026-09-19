/**
 * Compress Pictures (Word's Picture Format → Adjust → Compress Pictures):
 * re-encode the selected picture at a target resolution (96/150/220 ppi of
 * its display size, never upscaling) and optionally bake away the crop
 * window ("delete cropped areas"). The encoder runs through a canvas in the
 * renderer; the sizing math is pure so tests can drive it without jsdom
 * canvas support.
 */

export interface CompressCrop {
  /** source crop fractions (a:srcRect), each 0..1 */
  l: number
  t: number
  r: number
  b: number
}

export interface CompressOptions {
  /** target resolution: 96 = e-mail, 150 = web, 220 = print (Word presets) */
  ppi: 96 | 150 | 220
  /** bake the crop window into the new bytes and drop a:srcRect */
  deleteCropped: boolean
  /** current source crop; ignored unless deleteCropped */
  crop?: CompressCrop | null
}

/** the source window that survives compression (crop applied first when requested) */
export function compressSourceWindow(
  natural: { widthPx: number; heightPx: number },
  opts: CompressOptions,
): { x: number; y: number; widthPx: number; heightPx: number } {
  const w = Math.max(1, natural.widthPx)
  const h = Math.max(1, natural.heightPx)
  if (!opts.deleteCropped || !opts.crop) return { x: 0, y: 0, widthPx: w, heightPx: h }
  const frac = (v: unknown): number => (Number.isFinite(Number(v)) && Number(v) > 0 ? Number(v) : 0)
  // sides may not overlap: each is capped by what the opposite side left
  const l = Math.min(1, frac(opts.crop.l))
  const t = Math.min(1, frac(opts.crop.t))
  const r = Math.min(1 - l, frac(opts.crop.r))
  const b = Math.min(1 - t, frac(opts.crop.b))
  return {
    x: l * w,
    y: t * h,
    widthPx: Math.max(1, w * Math.max(0.01, 1 - l - r)),
    heightPx: Math.max(1, h * Math.max(0.01, 1 - t - b)),
  }
}

/**
 * Pixel size the compressed picture gets: the visible crop window at its
 * display size re-sampled to the chosen ppi (96 CSS px = 1 inch on the
 * page), capped at 1 (Word never enlarges), then applied to whatever source
 * area is encoded (full picture unless cropped areas are deleted).
 */
export function compressTargetSize(
  natural: { widthPx: number; heightPx: number },
  display: { widthPx: number; heightPx: number },
  opts: CompressOptions,
): { widthPx: number; heightPx: number } {
  const w = Math.max(1, natural.widthPx)
  // visible window (the part the page shows, crop or not); only the width
  // drives the scale — the height follows the encoded window's aspect
  let visibleW = w
  if (opts.crop) {
    const frac = (v: unknown): number =>
      Number.isFinite(Number(v)) && Number(v) > 0 ? Number(v) : 0
    const l = Math.min(1, frac(opts.crop.l))
    const r = Math.min(1 - l, frac(opts.crop.r))
    visibleW = Math.max(1, w * Math.max(0.01, 1 - l - r))
  }
  const dispW = Math.max(1, display.widthPx || visibleW)
  const scale = Math.min(1, ((dispW / 96) * opts.ppi) / visibleW)
  const window = compressSourceWindow(natural, opts)
  return {
    widthPx: Math.max(1, Math.round(window.widthPx * scale)),
    heightPx: Math.max(1, Math.round(window.heightPx * scale)),
  }
}

/**
 * Output type: PNG keeps transparency (crop/cutout results stay lossless in
 * shape); opaque sources re-encode as JPEG, which compresses photos far
 * smaller. GIF sources also go to PNG (no animation survives a static
 * re-encode anyway).
 */
export function compressedMime(srcMime: string): 'image/png' | 'image/jpeg' {
  return srcMime === 'image/jpeg' ? 'image/jpeg' : 'image/png'
}

export interface CompressResult {
  dataUrl: string
  widthPx: number
  heightPx: number
  mime: string
}

/** test seam: jsdom has no canvas, tests inject a fake encoder */
export type CompressEncoder = (
  dataUrl: string,
  window: { x: number; y: number; widthPx: number; heightPx: number },
  target: { widthPx: number; heightPx: number },
  mime: 'image/png' | 'image/jpeg',
) => Promise<string | null>

/** natural-size probe (test seam: jsdom images never decode) */
export type SizeProbe = (dataUrl: string) => Promise<{ widthPx: number; heightPx: number } | null>

/** default encoder: decode → draw the source window scaled → re-encode */
const canvasEncode: CompressEncoder = async (dataUrl, window, target, mime) => {
  const img = await new Promise<HTMLImageElement | null>((resolve) => {
    const el = new Image()
    el.onload = () => resolve(el)
    el.onerror = () => resolve(null)
    el.src = dataUrl
  })
  if (!img || !img.naturalWidth || !img.naturalHeight) return null
  const canvas = document.createElement('canvas')
  canvas.width = target.widthPx
  canvas.height = target.heightPx
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  ctx.drawImage(
    img,
    window.x,
    window.y,
    window.widthPx,
    window.heightPx,
    0,
    0,
    target.widthPx,
    target.heightPx,
  )
  return canvas.toDataURL(mime, 0.82)
}

/**
 * Compress a picture's bytes. Returns null when nothing would shrink (decode
 * failure, or the re-encode came out not smaller than the source).
 */
export async function compressPictureDataUrl(
  dataUrl: string,
  display: { widthPx: number; heightPx: number },
  opts: CompressOptions,
  encode: CompressEncoder = canvasEncode,
  probeSize: SizeProbe = imageSizeOf,
): Promise<CompressResult | null> {
  const mimeMatch = /^data:(image\/(?:png|jpeg|gif));base64,/.exec(dataUrl)
  if (!mimeMatch) return null
  const natural = await probeSize(dataUrl)
  if (!natural) return null
  const window = compressSourceWindow(natural, opts)
  const target = compressTargetSize(natural, display, opts)
  const mime = compressedMime(mimeMatch[1])
  const out = await encode(dataUrl, window, target, mime)
  if (!out) return null
  // only swap bytes that actually shrink (base64 length ≈ byte length)
  if (out.length >= dataUrl.length) return null
  return { dataUrl: out, widthPx: target.widthPx, heightPx: target.heightPx, mime }
}

/** natural size via a decode probe; null when the bytes do not decode */
export function imageSizeOf(
  dataUrl: string,
): Promise<{ widthPx: number; heightPx: number } | null> {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () =>
      resolve(
        img.naturalWidth && img.naturalHeight
          ? { widthPx: img.naturalWidth, heightPx: img.naturalHeight }
          : null,
      )
    img.onerror = () => resolve(null)
    img.src = dataUrl
  })
}

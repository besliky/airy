import { convertEmfToDataUrl, convertWmfToDataUrl } from './vendor/emf-converter/index.mjs'

const EMF_MIMES = new Set(['image/emf', 'image/x-emf'])
// localized GDI facenames mapped to their CSS-resolvable names
const FONT_FAMILY_MAP = {
  游ゴシック: 'Yu Gothic', // yuu goshikku
  游明朝: 'Yu Mincho', // yuu minchou
  メイリオ: 'Meiryo',
  'ｍｓ ｐゴシック': 'MS PGothic',
  'ｍｓ ゴシック': 'MS Gothic',
  'ｍｓ ｕｉゴシック': 'MS UI Gothic',
  'ｍｓ ｐ明朝': 'MS PMincho',
  'ｍｓ 明朝': 'MS Mincho',
}
const WMF_MIMES = new Set(['image/wmf', 'image/x-wmf'])
// gzip-compressed metafiles (.emz/.wmz)
const EMZ_MIMES = new Set(['image/emz', 'image/x-emz'])
const WMZ_MIMES = new Set(['image/wmz', 'image/x-wmz'])

export function isMetafileMime(mime: string | undefined): mime is string {
  return (
    mime !== undefined &&
    (EMF_MIMES.has(mime) || WMF_MIMES.has(mime) || EMZ_MIMES.has(mime) || WMZ_MIMES.has(mime))
  )
}

function isGzip(bytes: Uint8Array): boolean {
  return bytes.length > 2 && bytes[0] === 0x1f && bytes[1] === 0x8b
}

async function gunzip(bytes: Uint8Array): Promise<Uint8Array> {
  // copy to a fresh ArrayBuffer-backed view (BlobPart rejects ArrayBufferLike)
  const stream = new Blob([new Uint8Array(bytes)])
    .stream()
    .pipeThrough(new DecompressionStream('gzip'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

/** EMR_HEADER iType plus the ' EMF' signature at offset 40 */
function looksLikeEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 44) return false
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520
}

/** placeable-WMF magic, or a standard META_HEADER (type 1/2, HeaderSize 9) */
function looksLikeWmf(bytes: Uint8Array): boolean {
  if (bytes.length < 18) return false
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (dv.getUint32(0, true) === 0x9ac6cdd7) return true
  const type = dv.getUint16(0, true)
  return (type === 1 || type === 2) && dv.getUint16(2, true) === 9
}

// the engine compiles without DOM libs (it also runs under node); this scan is
// reached only in renderer environments, typed structurally (same as tiff.ts)
interface ImageBitmapLike {
  width: number
  height: number
  close(): void
}
interface OffscreenCanvasLike {
  getContext(id: '2d'): {
    drawImage(img: ImageBitmapLike, x: number, y: number): void
    getImageData(x: number, y: number, w: number, h: number): { data: Uint8ClampedArray }
  } | null
}
interface DomGlobals {
  createImageBitmap?(blob: Blob): Promise<ImageBitmapLike>
  OffscreenCanvas?: new (width: number, height: number) => OffscreenCanvasLike
}

/**
 * Decode a PNG data URL and scan its alpha channel. False means the frame is
 * fully transparent — a valid header with no drawable records converts to an
 * empty canvas. Null means the check cannot run (no canvas API — non-renderer
 * environments keep their converter result untouched).
 */
async function dataUrlHasPixels(dataUrl: string): Promise<boolean | null> {
  const dom = globalThis as DomGlobals
  if (!dom.createImageBitmap || !dom.OffscreenCanvas) return null
  try {
    const blob = await (await fetch(dataUrl)).blob()
    const bmp = await dom.createImageBitmap(blob)
    try {
      const cnv = new dom.OffscreenCanvas(bmp.width, bmp.height)
      const ctx = cnv.getContext('2d')
      if (!ctx) return null
      ctx.drawImage(bmp, 0, 0)
      const px = ctx.getImageData(0, 0, bmp.width, bmp.height).data
      for (let i = 3; i < px.length; i += 4) {
        if (px[i] !== 0) return true
      }
      return false
    } finally {
      try {
        bmp.close()
      } catch {
        // already detached
      }
    }
  } catch {
    return null
  }
}

/** kind name of a metafile mime for user-facing plates: EMF/EMZ/WMF/WMZ */
function metafileKind(mime: string): string {
  return mime.replace(/^image\/(?:x-)?/, '').toUpperCase()
}

/**
 * Plate text for a metafile part that cannot be shown ("WMF image (2 KB)"): the
 * broken-image frame otherwise gives no hint an image was there, nor how much
 * of the document it occupied (UX-1710).
 */
export function metafilePlateText(mime: string, byteLength: number): string {
  const kb = Math.max(1, Math.round(byteLength / 1024))
  return `${metafileKind(mime)} image (${kb} KB)`
}

/**
 * Render EMF/WMF (or gzipped EMZ/WMZ) bytes to a PNG data URL via the vendored
 * emf-converter. Returns null on parse failure, when no canvas API exists
 * (non-renderer environments), or when the render came out blank — callers keep
 * their existing empty-frame degrade. Failures are logged instead of silently
 * swallowed.
 */
export async function metafileToDataUrl(
  bytes: ArrayBuffer | Uint8Array,
  mime: string,
): Promise<string | null> {
  try {
    let u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    if (isGzip(u8)) u8 = await gunzip(u8)
    const buffer = u8.slice().buffer
    if (!isMetafileMime(mime)) return null
    // signature beats the declared mime: HWP-exported docx ship EMF bytes
    // under .wmf part names; mime only decides indeterminate bytes
    let isEmf: boolean
    if (looksLikeEmf(u8)) isEmf = true
    else if (looksLikeWmf(u8)) isEmf = false
    else isEmf = EMF_MIMES.has(mime) || EMZ_MIMES.has(mime)
    const opts = { dpiScale: 2, fontFamilyMap: FONT_FAMILY_MAP }
    const result = isEmf
      ? await convertEmfToDataUrl(buffer, opts)
      : await convertWmfToDataUrl(buffer, opts)
    if (result === null) {
      console.warn(`metafileToDataUrl: converter returned null (${mime}, ${u8.byteLength} bytes)`)
      return null
    }
    // UX-1710: a metafile whose header parses but that carries no drawable
    // records yields a fully transparent frame — a silent empty rectangle on
    // the canvas. Report it as a conversion failure so the caller's
    // unshowable-picture plate (with the part kind and size) shows instead.
    if ((await dataUrlHasPixels(result)) === false) {
      console.warn(`metafileToDataUrl: blank render (${mime}, ${u8.byteLength} bytes)`)
      return null
    }
    return result
  } catch (err) {
    console.warn(`metafileToDataUrl: conversion failed (${mime}):`, err)
    return null
  }
}

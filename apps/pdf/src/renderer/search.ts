import type { PDFDocumentProxy } from 'pdfjs-dist'

/** One hit: original page + PDF user-space rects (multiple when spanning several text items) */
export interface SearchMatch {
  pageIndex: number
  rects: [number, number, number, number][]
}

interface IndexedItem {
  start: number
  end: number
  x: number
  y: number
  w: number
  h: number
  /** Rotated run (tilted baseline) — excluded from block grouping */
  rot?: boolean
  /** pdf.js font id (e.g. 'g_d0_f7'); resolves to the run's font for edit previews */
  font?: string
}

export interface PageEntry {
  /** Original text (same length as lower; used for context excerpts) */
  text: string
  lower: string
  items: IndexedItem[]
}

export type SearchIndex = PageEntry[]

/** Upper bound on collected matches; the scan keeps counting hits past it so the UI can say "+N more" instead of silently truncating (BUG-1731) */
export const MAX_MATCHES = 1000

/** Full-text search outcome. When the cap stopped collection early, `capped` is
    true and `moreCount` holds the hits beyond the returned set (counted cheaply,
    without rect interpolation — occurrences with no item coverage are rare EOL
    gaps, so the count can overshoot by those). */
export interface SearchResult {
  matches: SearchMatch[]
  capped: boolean
  moreCount: number
}

/** Wrap-around step over a fixed match list (prev/next navigation) */
export function nextMatchIndex(current: number, dir: 1 | -1, count: number): number {
  return (current + dir + count) % count
}

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
  fontName?: string
}

/** Concatenate text per page + record each item's char range and PDF-space box (built once, cached per doc by caller) */
export async function buildSearchIndex(doc: PDFDocumentProxy): Promise<SearchIndex> {
  const entries: PageEntry[] = []
  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    let text = ''
    const items: IndexedItem[] = []
    for (const it of content.items as RawTextItem[]) {
      if (typeof it.str !== 'string') continue
      if (it.str.length > 0 && it.transform) {
        const h = it.height || Math.hypot(it.transform[2] ?? 0, it.transform[3] ?? 0)
        // Rotation tilts the baseline (b ≠ 0). A non-zero c alone is horizontal
        // shear — synthetic italics — which stays horizontally set and must keep
        // participating in block grouping.
        const rot = Math.abs(it.transform[1] ?? 0) > h * 1e-3
        items.push({
          start: text.length,
          end: text.length + it.str.length,
          x: it.transform[4] ?? 0,
          y: it.transform[5] ?? 0,
          w: it.width ?? 0,
          h,
          ...(rot ? { rot: true } : {}),
          ...(typeof it.fontName === 'string' ? { font: it.fontName } : {}),
        })
        text += it.str
      }
      if (it.hasEOL) text += '\n'
    }
    entries.push({ text, lower: text.toLowerCase(), items })
  }
  return entries
}

/** Case-insensitive full-text search; rects linearly interpolated within items by char ratio (approximate; bounding box for rotated glyphs).
    Collects at most MAX_MATCHES matches (first in document order) and keeps counting the remainder so callers can show an honest overflow. */
export function searchInIndex(index: SearchIndex, query: string): SearchResult {
  const q = query.toLowerCase()
  if (!q) return { matches: [], capped: false, moreCount: 0 }
  const matches: SearchMatch[] = []
  let capped = false
  let moreCount = 0
  for (let pageIndex = 0; pageIndex < index.length; pageIndex++) {
    const { lower, items } = index[pageIndex]!
    let from = 0
    for (;;) {
      const s = lower.indexOf(q, from)
      if (s < 0) break
      const e = s + q.length
      from = e
      if (matches.length >= MAX_MATCHES) {
        // Past the cap: count the hit without rect work (the expensive part) and move on
        capped = true
        moreCount++
        continue
      }
      const rects: [number, number, number, number][] = []
      for (const it of items) {
        if (it.end <= s || it.start >= e) continue
        const len = it.end - it.start
        const lo = (Math.max(s, it.start) - it.start) / len
        const hi = (Math.min(e, it.end) - it.start) / len
        const x1 = it.x + it.w * lo
        const x2 = it.x + it.w * hi
        if (x2 - x1 < 0.01) continue
        rects.push([x1, it.y, x2, it.y + it.h])
      }
      if (rects.length > 0) matches.push({ pageIndex, rects })
    }
  }
  return { matches, capped, moreCount }
}

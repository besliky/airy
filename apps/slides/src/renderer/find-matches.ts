/**
 * Find logic shared by the Find/Replace panel: flattens render-tree text
 * (mirroring the legacy layoutText traversal), locates hits under match-case /
 * whole-word conditions, and maps hit ranges back to run line boxes for the
 * canvas highlight overlay. Pure data in / data out — no React, no DOM.
 */
import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  RenderTextLayout,
  ShapeRenderNode,
  TableRenderNode,
} from '@airy-office/pptx-render'

export interface FindMatch {
  slideIndex: number
  sourceId: string
  /** hit range in the element's flattened text (same traversal as the segment walk) */
  start: number
  end: number
}

export interface FindConditions {
  matchCase: boolean
  wholeWord: boolean
}

/** one same-format span of the flattened node text; `rect` is its line box in
 *  top-node-local px (absent for synthetic separators and skip-highlight layouts) */
interface NodeSegment {
  text: string
  rect?: { x: number; y: number; w: number; h: number }
}

const isWordChar = (ch: string | undefined) => !!ch && /[\p{L}\p{N}_]/u.test(ch)

/** length-preserving lowercase: chars whose lowercase grows ('İ' → 'i̇') stay as-is so match offsets never shift */
export function foldCase(s: string): string {
  let out = ''
  for (const ch of s) {
    const lower = ch.toLowerCase()
    out += lower.length === ch.length ? lower : ch
  }
  return out
}

/** occurrences of needle in hay under the conditions, as [start, end) ranges */
export function hitRanges(hay: string, needle: string, c: FindConditions): Array<[number, number]> {
  const out: Array<[number, number]> = []
  if (!needle) return out
  const h = c.matchCase ? hay : foldCase(hay)
  const q = c.matchCase ? needle : foldCase(needle)
  let i = h.indexOf(q)
  while (i >= 0) {
    const end = i + q.length
    if (!c.wholeWord || (!isWordChar(hay[i - 1]) && !isWordChar(hay[end]))) {
      out.push([i, end])
      i = h.indexOf(q, end)
    } else {
      i = h.indexOf(q, i + 1)
    }
  }
  return out
}

/**
 * Flatten a laid-out text block into segments. `boxable: false` emits text
 * without rects — vertical layouts (lines are full columns) and WordArt warps
 * cannot be covered by a run line box; callers fall back to the element box.
 */
function layoutSegments(
  text: RenderTextLayout | undefined,
  ox: number,
  oy: number,
  boxable: boolean,
  out: NodeSegment[],
): void {
  if (!text) return
  const il = text.insets?.l ?? 0
  const it = text.insets?.t ?? 0
  text.lines.forEach((l, i) => {
    if (i > 0 && l.paraStart !== false) out.push({ text: '\n' })
    for (const r of l.runs) {
      if (r.isBullet) continue
      out.push({
        text: r.text,
        rect: boxable
          ? { x: ox + il + r.x, y: oy + it + l.top, w: r.widthPx, h: l.height }
          : undefined,
      })
    }
    if (l.trailingSpace) out.push({ text: l.trailingText ?? ' ' })
  })
}

/**
 * Flatten a node (group children / table cells included) into segments; rects
 * are relative to the top node's box origin. Flip parity follows the renderer:
 * a shape's own text layer is counter-flipped back to readable (identity
 * coords), table cells mirror under an effective flip, and a flipped group
 * mirrors its children's box origins within the group box.
 */
function nodeSegments(
  n: RenderNode,
  ox: number,
  oy: number,
  effH: boolean,
  effV: boolean,
  out: NodeSegment[],
): void {
  if (n.type === 'text' || n.type === 'shape') {
    const shape = n as ShapeRenderNode
    const boxable = !shape.text?.vert && !shape.text?.txWarp
    layoutSegments(shape.text, ox, oy, boxable, out)
    return
  }
  if (n.type === 'table') {
    const table = n as TableRenderNode
    table.cells.forEach((c, i) => {
      if (i > 0) out.push({ text: '\n' })
      const x = effH ? ox + n.box.w - (c.x + c.w) : ox + c.x
      const y = effV ? oy + n.box.h - (c.y + c.h) : oy + c.y
      layoutSegments(c.text, x, y, !c.text?.vert, out)
    })
    return
  }
  if (n.type === 'group') {
    const g = n as GroupRenderNode
    g.children.forEach((c, i) => {
      if (i > 0) out.push({ text: '\n' })
      const cx = effH ? n.box.w - (c.box.x + c.box.w) : c.box.x
      const cy = effV ? n.box.h - (c.box.y + c.box.h) : c.box.y
      nodeSegments(c, ox + cx, oy + cy, effH !== !!c.box.flipH, effV !== !!c.box.flipV, out)
    })
  }
}

/** the node's flattened searchable text */
export function nodeText(n: RenderNode): string {
  const segs: NodeSegment[] = []
  nodeSegments(n, 0, 0, !!n.box.flipH, !!n.box.flipV, segs)
  return segs.map((s) => s.text).join('')
}

/** rects covering [start, end) of the flattened text, in top-node-local px;
 *  empty when the layout cannot be boxed (vertical / warped → element fallback) */
export function matchRects(
  n: RenderNode,
  start: number,
  end: number,
): Array<{ x: number; y: number; w: number; h: number }> {
  const segs: NodeSegment[] = []
  nodeSegments(n, 0, 0, !!n.box.flipH, !!n.box.flipV, segs)
  const rects: Array<{ x: number; y: number; w: number; h: number }> = []
  let off = 0
  for (const s of segs) {
    const sStart = off
    off += s.text.length
    if (!s.rect || off <= start || sStart >= end) continue
    const a = Math.max(start, sStart) - sStart
    const b = Math.min(end, off) - sStart
    const len = s.text.length
    if (len <= 0) continue
    // partial-run coverage interpolates inside the run's box — exact at run
    // boundaries, approximate inside a variable-width run
    const f0 = a / len
    const f1 = b / len
    rects.push({
      x: s.rect.x + f0 * s.rect.w,
      y: s.rect.y,
      w: (f1 - f0) * s.rect.w,
      h: s.rect.h,
    })
  }
  return rects
}

/** all hits across the deck, one FindMatch per occurrence */
export function buildMatches(
  slides: RenderSlide[],
  query: string,
  conditions: FindConditions | boolean,
): FindMatch[] {
  const c: FindConditions =
    typeof conditions === 'boolean' ? { matchCase: conditions, wholeWord: false } : conditions
  const out: FindMatch[] = []
  if (!query) return out
  slides.forEach((sl, si) => {
    for (const n of sl.nodes) {
      if (n.decoration) continue
      for (const [start, end] of hitRanges(nodeText(n), query, c)) {
        out.push({ slideIndex: si, sourceId: n.sourceId, start, end })
      }
    }
  })
  return out
}

/** find a node (recursing into groups) and return it with its box origin in slide px */
export function findNodeBox(
  slide: RenderSlide | undefined,
  sourceId: string,
): { node: RenderNode; x: number; y: number } | null {
  const walk = (
    n: RenderNode,
    ox: number,
    oy: number,
  ): { node: RenderNode; x: number; y: number } | null => {
    if (n.sourceId === sourceId) return { node: n, x: ox + n.box.x, y: oy + n.box.y }
    if (n.type === 'group') {
      for (const c of (n as GroupRenderNode).children) {
        const hit = walk(c, ox + n.box.x, oy + n.box.y)
        if (hit) return hit
      }
    }
    return null
  }
  for (const n of slide?.nodes ?? []) {
    const hit = walk(n, 0, 0)
    if (hit) return hit
  }
  return null
}

/** slide-px bounding box of a hit (union of its run boxes, or the element box when unboxable) */
export function matchFocusBox(
  slide: RenderSlide | undefined,
  m: FindMatch,
): { x: number; y: number; w: number; h: number } | undefined {
  const hit = findNodeBox(slide, m.sourceId)
  if (!hit) return undefined
  const rects = matchRects(hit.node, m.start, m.end)
  if (!rects.length) return { x: hit.x, y: hit.y, w: hit.node.box.w, h: hit.node.box.h }
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const r of rects) {
    minX = Math.min(minX, hit.x + r.x)
    minY = Math.min(minY, hit.y + r.y)
    maxX = Math.max(maxX, hit.x + r.x + r.w)
    maxY = Math.max(maxY, hit.y + r.y + r.h)
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY }
}

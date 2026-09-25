/** Column-aware selection correction for the pdf.js text layer (BUG-1729).
 *
 *  pdf.js TextLayer emits absolutely-positioned spans in content-stream order, and a
 *  native browser selection is a DOM-order interval between its two endpoints. On a
 *  multi-column page that order is column-major, so a narrow diagonal drag from the
 *  top of the left column into the right column selects the ENTIRE left column on the
 *  way through the DOM. Reordering the DOM could fix the drag, but would break
 *  reading-order copy of whole columns/documents (the two orderings are opposites),
 *  and a full custom mouse-driven selection engine would endanger single-column
 *  behavior. So instead: keep the native selection untouched everywhere it works and
 *  only CORRECT it after the fact, on pages where strict column detection succeeds.
 *
 *  Correction = the geometric "band" between the selection's two endpoints (the lines
 *  the anchor→focus segment sweeps through, Acrobat-style), applied as reading-ordered
 *  per-line ranges. Pages with no strict column layout (single column, full-width
 *  headings crossing the gutter, rotated runs) are never touched, so normal
 *  single-column selection and copy-paste behave exactly as before.
 *
 *  Everything above the DOM adapter is pure geometry over plain boxes so the detector
 *  and the band planner are unit-testable without real layout. */

export interface SpanBox {
  left: number
  top: number
  right: number
  bottom: number
}

export interface Point {
  x: number
  y: number
}

/** A detected column layout: interior vertical gutters no line box crosses. */
export interface ColumnLayout {
  /** Sorted left to right; each gutter is the free x-interval between two columns */
  gutters: Array<{ left: number; right: number }>
}

/** A visual line: the union box of its member items (DOM spans in the adapter). */
export interface LineItem<T> {
  box: SpanBox
  items: T[]
}

/** A column cell: the part of one visual line inside one column (spans never cross a
    detected gutter, so cells are contiguous runs of a line's left-to-right items). */
export interface CellItem<T> {
  box: SpanBox
  items: T[]
  /** Owning line (row), for band-vs-selection bookkeeping */
  line: LineItem<T>
}

/** A column layout is only trusted with enough evidence; fewer visual lines than this
    cannot distinguish a two-column body from loose single-column text. */
const MIN_LINES_FOR_COLUMNS = 6

/** A line is "selected" when at least this fraction of its area intersects a
    selection client rect (end-of-band lines are typically only partially covered). */
const LINE_SELECTED_AREA = 0.3

/** A line belongs to the band when its box comes within this fraction of the line's
    height of the anchor→focus segment (the swept band is ~one line tall around it). */
const BAND_TOLERANCE = 0.75

/** Minimum gutter width, relative to the median line height / total text width:
    wider than any word gap inside a single column could be. */
const GUTTER_OF_LINE_HEIGHT = 0.8
const GUTTER_OF_TEXT_WIDTH = 0.08

/** Correction only fires on clear over-reach: the native selection must contain at
    least this many lines beyond the computed band, else the DOM order is close enough
    to the visual band that rewriting would only churn the selection. */
const MIN_OVERREACH_LINES = 3

/** Each column must hold at least this fraction of the page's lines (and at least 2). */
const MIN_COLUMN_LINE_SHARE = 0.2

const boxHeight = (b: SpanBox) => b.bottom - b.top

/** Group items into visual lines: boxes whose vertical overlap is at least half the
    smaller height share a line (same rule as text-line.ts, but page-wide). Returns
    lines top to bottom, member items sorted left to right. */
export function clusterLineBoxes<T>(items: T[], boxOf: (item: T) => SpanBox): LineItem<T>[] {
  const sorted = [...items].sort(
    (a, b) => boxOf(a).top - boxOf(b).top || boxOf(a).left - boxOf(b).left,
  )
  const lines: LineItem<T>[] = []
  for (const item of sorted) {
    const box = boxOf(item)
    const line = lines.find((l) => {
      const overlap = Math.min(box.bottom, l.box.bottom) - Math.max(box.top, l.box.top)
      return overlap >= Math.min(boxHeight(box), boxHeight(l.box)) * 0.5
    })
    if (line) {
      line.items.push(item)
      line.box = {
        left: Math.min(line.box.left, box.left),
        top: Math.min(line.box.top, box.top),
        right: Math.max(line.box.right, box.right),
        bottom: Math.max(line.box.bottom, box.bottom),
      }
    } else {
      lines.push({ box: { ...box }, items: [item] })
    }
  }
  for (const line of lines) line.items.sort((a, b) => boxOf(a).left - boxOf(b).left)
  return lines
}

/** Free x-gaps between a line's own boxes (its internal gaps, any width). */
const lineGaps = <T>(line: LineItem<T>, boxOf: (item: T) => SpanBox) => {
  const boxes = line.items.map(boxOf).sort((a, b) => a.left - b.left)
  const gaps: Array<{ left: number; right: number }> = []
  for (let i = 1; i < boxes.length; i++)
    gaps.push({ left: boxes[i - 1]!.right, right: boxes[i]!.left })
  return gaps
}

/** Detect a strict multi-column layout: interior vertical gutters that no line box
    crosses, wide enough to be a gutter (not a word gap), with every resulting column
    holding a meaningful share of the lines. Any full-width element (title rule,
    figure) crosses the gutter and makes detection fail — deliberately conservative,
    those pages keep the plain native selection. Null for single-column pages. */
export function detectColumns<T>(items: T[], boxOf: (item: T) => SpanBox): ColumnLayout | null {
  const lines = clusterLineBoxes(items, boxOf)
  if (lines.length < MIN_LINES_FOR_COLUMNS) return null

  const heights = lines.map((l) => boxHeight(l.box)).sort((a, b) => a - b)
  const medianHeight = heights[Math.floor(heights.length / 2)]!
  if (!(medianHeight > 0)) return null
  const textLeft = Math.min(...lines.map((l) => l.box.left))
  const textRight = Math.max(...lines.map((l) => l.box.right))
  const textWidth = textRight - textLeft
  if (!(textWidth > 0)) return null
  const minGutter = Math.max(medianHeight * GUTTER_OF_LINE_HEIGHT, textWidth * GUTTER_OF_TEXT_WIDTH)

  // Candidate gutters start as the wider internal gaps of the first line, then get
  // shrunk by intersecting with every other line's gaps. A candidate survives only if
  // every line has a gap covering it (no box crosses the gutter) and the surviving
  // interval is still gutter-wide; intervals only ever shrink, so surviving width is
  // monotone and an early exit is safe.
  const candidates = lineGaps(lines[0]!, boxOf).filter((g) => g.right - g.left >= minGutter)
  const found: Array<{ left: number; right: number }> = []
  for (const candidate of candidates) {
    let lo = candidate.left
    let hi = candidate.right
    let alive = true
    for (let i = 1; alive && i < lines.length; i++) {
      let best: { left: number; right: number } | null = null
      for (const gap of lineGaps(lines[i]!, boxOf)) {
        const left = Math.max(lo, gap.left)
        const right = Math.min(hi, gap.right)
        if (right >= left && (!best || right - left > best.right - best.left))
          best = { left, right }
      }
      if (!best || best.right - best.left < minGutter) alive = false
      else {
        lo = best.left
        hi = best.right
      }
    }
    if (alive) found.push({ left: lo, right: hi })
  }
  // Overlapping candidates converge to the same interval while shrinking
  const gutters = found
    .sort((a, b) => a.left - b.left)
    .filter((g, i, all) => i === 0 || g.left >= all[i - 1]!.right)
  if (gutters.length === 0) return null

  // Every resulting column must hold a meaningful share of the page's lines
  const perColumn = new Map<number, number>()
  for (const line of lines) {
    const col = columnOfBox(line.box, { gutters })
    perColumn.set(col, (perColumn.get(col) ?? 0) + 1)
  }
  const minColumnLines = Math.max(2, Math.ceil(lines.length * MIN_COLUMN_LINE_SHARE))
  if ([...perColumn.values()].some((n) => n < minColumnLines)) return null
  return { gutters }
}

/** Column index of a box: how many gutters lie fully to the box center's left. */
export function columnOfBox(box: SpanBox, layout: ColumnLayout): number {
  return columnIndexOf({ x: (box.left + box.right) / 2, y: box.top }, layout)
}

/** Column index of a horizontal position: how many gutters lie fully to its left. */
export function columnIndexOf(point: Point, layout: ColumnLayout): number {
  let col = 0
  for (const g of layout.gutters) if (point.x >= g.right) col++
  return col
}

/** Distance between a point and a segment. */
function distPointToSegment(p: Point, a: Point, b: Point): number {
  const dx = b.x - a.x
  const dy = b.y - a.y
  const len2 = dx * dx + dy * dy
  const t = len2 === 0 ? 0 : Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / len2))
  return Math.hypot(a.x + t * dx - p.x, a.y + t * dy - p.y)
}

/** Distance from a point to a box (0 inside). */
const distPointToBox = (p: Point, b: SpanBox) =>
  Math.hypot(Math.max(b.left - p.x, 0, p.x - b.right), Math.max(b.top - p.y, 0, p.y - b.bottom))

const boxCorners = (b: SpanBox): Point[] => [
  { x: b.left, y: b.top },
  { x: b.right, y: b.top },
  { x: b.right, y: b.bottom },
  { x: b.left, y: b.bottom },
]

/** Segment (a→b) intersects a box. */
function segmentIntersectsBox(a: Point, b: Point, box: SpanBox): boolean {
  if (distPointToBox(a, box) === 0 || distPointToBox(b, box) === 0) return true
  const corners = boxCorners(box)
  const orient = (p: Point, q: Point, r: Point) =>
    Math.sign((q.x - p.x) * (r.y - p.y) - (q.y - p.y) * (r.x - p.x))
  const segCross = (p: Point, q: Point, r: Point, s: Point) =>
    orient(p, q, r) !== orient(p, q, s) && orient(r, s, p) !== orient(r, s, q)
  for (let i = 0; i < 4; i++) if (segCross(a, b, corners[i]!, corners[(i + 1) % 4])) return true
  return false
}

/** Distance from a line box to the anchor→focus segment: 0 when they intersect,
    otherwise the exact minimum (for two convex sets it is attained at a rect corner
    vs the segment, or a segment endpoint vs the rect). Used to keep the selection
    band ~one line tall around the swept path. */
export function distBoxToSegment(box: SpanBox, a: Point, b: Point): number {
  if (segmentIntersectsBox(a, b, box)) return 0
  let best = Math.min(distPointToBox(a, box), distPointToBox(b, box))
  for (const c of boxCorners(box)) best = Math.min(best, distPointToSegment(c, a, b))
  return best
}

/** Split a visual line into per-column cells (a line on a single-column page is one
    cell; on a multi-column page its items never cross a detected gutter). */
export function splitLineCells<T>(
  line: LineItem<T>,
  layout: ColumnLayout,
  boxOf: (item: T) => SpanBox,
): CellItem<T>[] {
  const cells: CellItem<T>[] = []
  for (const item of line.items) {
    const box = boxOf(item)
    const col = columnOfBox(box, layout)
    const last = cells[cells.length - 1]
    const lastBox = last ? boxOf(last.items[last.items.length - 1]!) : null
    if (last && lastBox && columnOfBox(lastBox, layout) === col) {
      last.items.push(item)
      last.box.left = Math.min(last.box.left, box.left)
      last.box.top = Math.min(last.box.top, box.top)
      last.box.right = Math.max(last.box.right, box.right)
      last.box.bottom = Math.max(last.box.bottom, box.bottom)
    } else {
      cells.push({ box: { ...box }, items: [item], line })
    }
  }
  return cells
}

/** The visual band between the selection endpoints: column cells whose boxes the
    anchor→focus segment sweeps through, ordered for reading (column by column, top to
    bottom — Acrobat's copy order for a band across columns). Null when both endpoints
    share a column — the native DOM-order selection is already the reading order there,
    so it is left alone. */
export function computeSelectionBand<T>(
  lines: LineItem<T>[],
  layout: ColumnLayout,
  anchor: Point,
  focus: Point,
  boxOf: (item: T) => SpanBox,
): CellItem<T>[] | null {
  if (columnIndexOf(anchor, layout) === columnIndexOf(focus, layout)) return null
  const cells = lines
    .flatMap((line) => splitLineCells(line, layout, boxOf))
    .filter(
      (c) => distBoxToSegment(c.box, anchor, focus) <= BAND_TOLERANCE * (c.box.bottom - c.box.top),
    )
  if (cells.length === 0) return null
  return cells.sort(
    (a, b) =>
      columnOfBox(a.box, layout) - columnOfBox(b.box, layout) ||
      a.box.top - b.box.top ||
      a.box.left - b.box.left,
  )
}

/** Is a line covered by the given selection client rects? Coverage is summed over
    rects (selection paints one rect per span, so a full line is covered piecewise). */
const lineIsSelected = (box: SpanBox, rects: SpanBox[]) => {
  const area = Math.max(1, (box.right - box.left) * (box.bottom - box.top))
  let covered = 0
  for (const r of rects) {
    const w = Math.min(box.right, r.right) - Math.max(box.left, r.left)
    const h = Math.min(box.bottom, r.bottom) - Math.max(box.top, r.top)
    if (w > 0 && h > 0) covered += w * h
  }
  return covered / area >= LINE_SELECTED_AREA
}

/** Decision layer: should the current native selection be rewritten to the band?
    Returns the corrected column cells in reading order (each with the items its range
    should span), or null to keep the native selection as is. Fires only on clear
    over-reach — the native selection covering whole extra rows the band does not
    want — and never on plain select-all. */
export function planBandCorrection<T>(params: {
  layout: ColumnLayout | null
  lines: LineItem<T>[]
  selectionRects: SpanBox[]
  anchor: Point
  focus: Point
  boxOf: (item: T) => SpanBox
}): CellItem<T>[] | null {
  const { layout, lines, selectionRects, anchor, focus, boxOf } = params
  if (!layout || lines.length === 0 || selectionRects.length === 0) return null
  const selectedRows = lines.filter((l) => lineIsSelected(l.box, selectionRects))
  if (selectedRows.length === 0) return null
  const band = computeSelectionBand(lines, layout, anchor, focus, boxOf)
  if (!band) return null
  // Row counts, not cells: the band covers each of its rows in every swept column,
  // and select-all / near-select-all are legitimate native selections, not over-reach
  const bandRows = new Set(band.map((c) => c.line)).size
  if (bandRows >= lines.length) return null
  if (selectedRows.length - bandRows < MIN_OVERREACH_LINES) return null
  return band
}

// ---------------------------------------------------------------------------
// DOM adapter: watches document selection changes and applies the correction.
// Thin on purpose — jsdom has no layout, so this half stays covered by review and
// the pure geometry tests above it.
// ---------------------------------------------------------------------------

interface LayerGeometry {
  spanCount: number
  layout: ColumnLayout | null
  lines: LineItem<SpanEl>[]
}

/** A measured text-layer span: element plus its client box at measure time. */
interface SpanEl {
  el: HTMLElement
  box: SpanBox
}

const geometryCache = new WeakMap<HTMLElement, LayerGeometry>()

/** pdf.js 6 rotates runs via an inline --rotate custom property (same heuristic as
    text-line.ts); rotated spans have no trustworthy axis-aligned box */
const isRotated = (el: HTMLElement) => {
  const r = el.style.getPropertyValue('--rotate').trim()
  return (r !== '' && r !== '0deg') || /rotate\(/.test(el.style.transform)
}

/** The .textLayer owning a range boundary node (the layer itself counts). */
const layerOfNode = (node: Node | null): HTMLElement | null => {
  const el = node?.nodeType === Node.ELEMENT_NODE ? (node as Element) : node?.parentElement
  return el?.closest('.textLayer') ?? null
}

/** Measure one page's text layer: span boxes, line clustering, column layout.
    Cached per layer element — pdf.js rebuilds the layer (new element) on zoom,
    rotation and re-render, which naturally invalidates the geometry. */
const measureLayer = (layer: HTMLElement): LayerGeometry | null => {
  const spans = [...layer.querySelectorAll<HTMLElement>('span')].filter(
    (el) => (el.textContent ?? '').trim() !== '' && !el.querySelector('span') && !isRotated(el),
  )
  const cached = geometryCache.get(layer)
  if (cached && cached.spanCount === spans.length) return cached
  const boxes = spans.map((el) => {
    const r = el.getBoundingClientRect()
    return { el, box: { left: r.left, top: r.top, right: r.right, bottom: r.bottom } }
  })
  const layout = detectColumns(boxes, (b) => b.box)
  const lines = clusterLineBoxes(boxes, (b) => b.box)
  const geometry: LayerGeometry = { spanCount: spans.length, layout, lines }
  geometryCache.set(layer, geometry)
  return geometry
}

/** Max valid boundary offset for a node (child index for elements, text length
    otherwise). */
const maxOffsetOf = (node: Node) =>
  node.nodeType === Node.ELEMENT_NODE ? node.childNodes.length : (node.textContent ?? '').length

/** Client point at a selection boundary (node, offset): the caret rect at that
    position, widened to one child/character when collapsed rects are unavailable. */
const pointAtBoundary = (
  node: Node | null,
  offset: number,
): { point: Point; box: SpanBox } | null => {
  if (!node || !layerOfNode(node)) return null
  const maxOffset = maxOffsetOf(node)
  const range = document.createRange()
  try {
    range.setStart(node, Math.min(offset, maxOffset))
    range.setEnd(node, Math.min(offset, maxOffset))
  } catch {
    return null
  }
  const unionOf = (r: Range): SpanBox | null => {
    let box: SpanBox | null = null
    for (const q of r.getClientRects()) {
      if (q.width <= 0 && q.height <= 0) continue
      box = box
        ? {
            left: Math.min(box.left, q.left),
            top: Math.min(box.top, q.top),
            right: Math.max(box.right, q.right),
            bottom: Math.max(box.bottom, q.bottom),
          }
        : { left: q.left, top: q.top, right: q.right, bottom: q.bottom }
    }
    return box
  }
  let box = unionOf(range)
  if (!box) {
    // Collapsed boundary rects are unreliable in some engines: widen by one unit
    try {
      range.setEnd(node, Math.min(offset + 1, maxOffset))
      box = unionOf(range)
    } catch {
      return null
    }
  }
  if (!box) return null
  return { point: { x: box.left, y: (box.top + box.bottom) / 2 }, box }
}

/** Range union box (for comparing an applied selection against the band). */
const rangeBox = (range: Range): SpanBox | null => {
  let box: SpanBox | null = null
  for (const r of range.getClientRects()) {
    if (r.width <= 0 && r.height <= 0) continue
    box = box
      ? {
          left: Math.min(box.left, r.left),
          top: Math.min(box.top, r.top),
          right: Math.max(box.right, r.right),
          bottom: Math.max(box.bottom, r.bottom),
        }
      : { left: r.left, top: r.top, right: r.right, bottom: r.bottom }
  }
  return box
}

const sameBox = (a: SpanBox, b: SpanBox) =>
  Math.abs(a.left - b.left) <= 1.5 &&
  Math.abs(a.top - b.top) <= 1.5 &&
  Math.abs(a.right - b.right) <= 1.5 &&
  Math.abs(a.bottom - b.bottom) <= 1.5

/** Apply the band as one range per column cell, in reading order. Falls back to a
    single DOM-order range when the engine does not support multi-range selections
    (that is today's behavior, so the fallback never gets worse than the bug). */
const applyBand = (selection: Selection, band: CellItem<SpanEl>[]) => {
  const spanOf = (cell: CellItem<SpanEl>) => ({
    first: cell.items[0]!.el,
    last: cell.items[cell.items.length - 1]!.el,
  })
  selection.removeAllRanges()
  try {
    for (const cell of band) {
      const { first, last } = spanOf(cell)
      const range = document.createRange()
      range.setStartBefore(first)
      range.setEndAfter(last)
      selection.addRange(range)
    }
  } catch {
    /* range construction over detached nodes: keep whatever applied */
  }
  if (selection.rangeCount < band.length) {
    // Multi-range unsupported: degrade to a single span of the band (previous behavior)
    const first = spanOf(band[0]!).first
    const last = spanOf(band[band.length - 1]!).last
    selection.removeAllRanges()
    try {
      const range = document.createRange()
      range.setStartBefore(first)
      range.setEndAfter(last)
      selection.addRange(range)
    } catch {
      /* ignore */
    }
  }
}

let scheduled = false

const correctSelection = () => {
  scheduled = false
  const selection = window.getSelection()
  // Only a single native range is ever corrected: a multi-range selection is either a
  // band this module applied (stable) or a selection another surface owns.
  if (!selection || selection.rangeCount !== 1 || selection.isCollapsed) return
  const range = selection.getRangeAt(0)
  const layer = layerOfNode(range.startContainer)
  if (!layer || layer !== layerOfNode(range.endContainer)) return
  const geometry = measureLayer(layer)
  if (!geometry?.layout) return

  const anchor = pointAtBoundary(selection.anchorNode, selection.anchorOffset)
  const focus = pointAtBoundary(selection.focusNode, selection.focusOffset)
  if (!anchor || !focus) return
  // Point each endpoint inward along the drag: the anchor's inner edge is the side
  // facing the focus (and vice versa), so the band segment follows the true sweep
  const focusRight = focus.point.x >= anchor.point.x
  const anchorPoint = {
    x: focusRight ? anchor.box.right : anchor.box.left,
    y: (anchor.box.top + anchor.box.bottom) / 2,
  }
  const focusPoint = {
    x: focusRight ? focus.box.left : focus.box.right,
    y: (focus.box.top + focus.box.bottom) / 2,
  }

  const selectionRects: SpanBox[] = [...range.getClientRects()]
    .filter((r) => r.width > 0 && r.height > 0)
    .map((r) => ({ left: r.left, top: r.top, right: r.right, bottom: r.bottom }))
  const band = planBandCorrection({
    layout: geometry.layout,
    lines: geometry.lines,
    selectionRects,
    anchor: anchorPoint,
    focus: focusPoint,
    boxOf: (spanEl) => spanEl.box,
  })
  if (!band) return
  // Already applied (idempotence): comparing geometry avoids correction loops, since
  // rewriting the selection fires selectionchange again
  if (
    selection.rangeCount === band.length &&
    band.every((line, i) => {
      const box = rangeBox(selection.getRangeAt(i))
      return box !== null && sameBox(box, line.box)
    })
  )
    return
  applyBand(selection, band)
}

/** Watch native selections and correct cross-column over-reach on text layers.
    `isEnabled` gates the correction (off while text editing owns the layer).
    Returns the disposer. */
export function watchColumnSelection(isEnabled: () => boolean): () => void {
  const onSelectionChange = () => {
    if (!isEnabled() || scheduled) return
    scheduled = true
    requestAnimationFrame(correctSelection)
  }
  document.addEventListener('selectionchange', onSelectionChange)
  return () => {
    document.removeEventListener('selectionchange', onSelectionChange)
    if (scheduled) scheduled = false
  }
}

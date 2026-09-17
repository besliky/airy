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
 *  top-node-local px (absent for synthetic separators and skip-highlight layouts);
 *  `map` carries the rect's ancestor rotation chain (top-node-local frame) */
interface NodeSegment {
  text: string
  rect?: { x: number; y: number; w: number; h: number }
  map?: AffineMap
}

/** 2D affine map p → (a·x + c·y + e, b·x + d·y + f); compositions of
 *  translations and clockwise (screen-coords) rotations stay det=1 */
export interface AffineMap {
  a: number
  b: number
  c: number
  d: number
  e: number
  f: number
}

const IDENTITY: AffineMap = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }

/** m ∘ n — apply n first, then m */
function composeAffine(m: AffineMap, n: AffineMap): AffineMap {
  return {
    a: m.a * n.a + m.c * n.b,
    b: m.b * n.a + m.d * n.b,
    c: m.a * n.c + m.c * n.d,
    d: m.b * n.c + m.d * n.d,
    e: m.a * n.e + m.c * n.f + m.e,
    f: m.b * n.e + m.d * n.f + m.f,
  }
}

function translationAffine(tx: number, ty: number): AffineMap {
  return { a: 1, b: 0, c: 0, d: 1, e: tx, f: ty }
}

function scaleAffine(sx: number, sy: number): AffineMap {
  return { a: sx, b: 0, c: 0, d: sy, e: 0, f: 0 }
}

/** clockwise rotation about (cx, cy) by deg — y-down screen coords (Konva/SVG parity) */
function rotationAffine(cx: number, cy: number, deg: number): AffineMap {
  const rad = (deg * Math.PI) / 180
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)
  return {
    a: cos,
    b: sin,
    c: -sin,
    d: cos,
    e: cx - cos * cx + sin * cy,
    f: cy - sin * cx - cos * cy,
  }
}

/** A node's local frame → its parent frame, Konva-exactly: boxPivotProps
 *  places each node with position=center, offset=center, rotation, and
 *  scaleX/scaleY = ±1 for its own flip, i.e.
 *  T(box.x, box.y) ∘ [T(center) ∘ R ∘ S ∘ T(−center)]. Composing these maps
 *  down the tree is what the nested Konva containers apply, so an ancestor
 *  flip arrives as a scale(−1) matrix — conjugating every descendant rotation
 *  it wraps (a mirrored group reverses its children's rotation direction). */
function nodeLocalToParentMap(box: {
  x: number
  y: number
  w: number
  h: number
  rotationDeg: number
  flipH?: boolean
  flipV?: boolean
}): AffineMap {
  const cx = box.w / 2
  const cy = box.h / 2
  // T(center) ∘ R ∘ S ∘ T(−center) — flip mirrors about the box center
  const pivotTransform = composeAffine(
    translationAffine(cx, cy),
    composeAffine(
      rotationAffine(0, 0, box.rotationDeg || 0),
      composeAffine(
        scaleAffine(box.flipH ? -1 : 1, box.flipV ? -1 : 1),
        translationAffine(-cx, -cy),
      ),
    ),
  )
  return composeAffine(translationAffine(box.x, box.y), pivotTransform)
}

/** apply the map to a point */
export function applyAffine(m: AffineMap, x: number, y: number): { x: number; y: number } {
  return { x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f }
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
  m: AffineMap,
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
        map: m,
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
  m: AffineMap,
  out: NodeSegment[],
): void {
  if (n.type === 'text' || n.type === 'shape') {
    const shape = n as ShapeRenderNode
    const boxable = !shape.text?.vert && !shape.text?.txWarp
    layoutSegments(shape.text, ox, oy, boxable, m, out)
    return
  }
  if (n.type === 'table') {
    const table = n as TableRenderNode
    table.cells.forEach((c, i) => {
      if (i > 0) out.push({ text: '\n' })
      const x = effH ? ox + n.box.w - (c.x + c.w) : ox + c.x
      const y = effV ? oy + n.box.h - (c.y + c.h) : oy + c.y
      layoutSegments(c.text, x, y, !c.text?.vert, m, out)
    })
    return
  }
  if (n.type === 'group') {
    const g = n as GroupRenderNode
    g.children.forEach((c, i) => {
      if (i > 0) out.push({ text: '\n' })
      const cx = effH ? n.box.w - (c.box.x + c.box.w) : c.box.x
      const cy = effV ? n.box.h - (c.box.y + c.box.h) : c.box.y
      // rect offsets already carry the (flip-mirrored) child origin through
      // ox/oy — the map contributes only the child's rotation, pivoting the
      // MIRRORED child center in this group's frame (boxPivotProps
      // semantics). Under an effective mirror (exactly one of effH/effV) the
      // rotation must be conjugated — a mirrored group reverses its
      // children's rotation direction, like Konva's scale(-1) container does.
      const rotationDeg =
        effH !== effV ? -((c.box.rotationDeg || 0) as number) : ((c.box.rotationDeg || 0) as number)
      const childMap = composeAffine(
        m,
        rotationAffine(cx + c.box.w / 2, cy + c.box.h / 2, rotationDeg),
      )
      nodeSegments(
        c,
        ox + cx,
        oy + cy,
        effH !== !!c.box.flipH,
        effV !== !!c.box.flipV,
        childMap,
        out,
      )
    })
  }
}

/** the node's flattened searchable text */
export function nodeText(n: RenderNode): string {
  const segs: NodeSegment[] = []
  nodeSegments(n, 0, 0, !!n.box.flipH, !!n.box.flipV, IDENTITY, segs)
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
  nodeSegments(n, 0, 0, !!n.box.flipH, !!n.box.flipV, IDENTITY, segs)
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

/** next match index for a find-next/find-prev step: without a cursor the
 *  first press enters at the first (next) or last (prev) match, otherwise the
 *  index wraps modulo the match count. Count 0 returns -1 (no matches). */
export function stepMatchIndex(cursor: number, dir: 1 | -1, count: number): number {
  if (count <= 0) return -1
  if (cursor < 0) return dir === 1 ? 0 : count - 1
  return (cursor + dir + count) % count
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

/** find a node (recursing into groups) and return it with its box origin in
 *  slide px plus the affine from the node's local frame (origin = its box
 *  top-left) to slide px, composing the ancestor-group translations and
 *  rotations AND flips — each node contributes its Konva-exact
 *  T(origin)∘T(center)∘R∘S∘T(−center) map (see nodeLocalToParentMap), so an
 *  ancestor flip rides along as scale(−1) and conjugates every descendant
 *  rotation it wraps. Flip parity stays positional only in the run-rect
 *  coordinate walk (nodeSegments), which conjugates child rotations itself
 *  for mirrors inside the found node's subtree. */
export function findNodeBox(
  slide: RenderSlide | undefined,
  sourceId: string,
): { node: RenderNode; x: number; y: number; map: AffineMap } | null {
  const walk = (
    n: RenderNode,
    ox: number,
    oy: number,
    m: AffineMap,
  ): { node: RenderNode; x: number; y: number; map: AffineMap } | null => {
    const local = composeAffine(m, nodeLocalToParentMap(n.box))
    if (n.sourceId === sourceId) return { node: n, x: ox + n.box.x, y: oy + n.box.y, map: local }
    if (n.type === 'group') {
      for (const c of (n as GroupRenderNode).children) {
        const hit = walk(c, ox + n.box.x, oy + n.box.y, local)
        if (hit) return hit
      }
    }
    return null
  }
  for (const n of slide?.nodes ?? []) {
    const hit = walk(n, 0, 0, IDENTITY)
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

/** an overlay rect in slide px: untransformed position + size, plus the
 *  rotation (degrees) and CSS-style transform-origin (px, relative to the
 *  rect's top-left) that reproduce the composed ancestor rotation chain */
export interface StageRect {
  x: number
  y: number
  w: number
  h: number
  rotation: number
  originX: number
  originY: number
}

/** axis-aligned local rect under an affine (rotation, possibly with a flip
 *  scale) → StageRect. With A = R(θ) + t, placing the element at r + t and
 *  rotating about the origin t (absolute stage px — i.e. transform-origin
 * −r, element-local) maps every corner q to exactly A·q; θ ≈ 0 collapses to
 * the plain rect. A det < 0 chain (a flipped ancestor contributes
 * scale(−1)) is first conjugated with a mirror about the RECT's own vertical
 * center axis — that mirror maps the rect onto itself, so the corner set is
 * unchanged while the map becomes det=1 and decomposes like any rotation. */
function projectRect(m: AffineMap, r: { x: number; y: number; w: number; h: number }): StageRect {
  let map = m
  if (map.a * map.d - map.b * map.c < 0) {
    // mirror x ↦ 2·(r.x + r.w/2) − x about the rect's center axis
    map = composeAffine(map, {
      a: -1,
      b: 0,
      c: 0,
      d: 1,
      e: 2 * (r.x + r.w / 2),
      f: 0,
    })
  }
  let theta = (Math.atan2(map.b, map.a) * 180) / Math.PI
  if (Math.abs(theta) < 1e-6) theta = 0
  if (theta === 0) {
    return { x: r.x + map.e, y: r.y + map.f, w: r.w, h: r.h, rotation: 0, originX: 0, originY: 0 }
  }
  return {
    x: r.x + map.e,
    y: r.y + map.f,
    w: r.w,
    h: r.h,
    rotation: theta,
    // avoid −0 (deep-equality and CSS-string parity)
    originX: r.x === 0 ? 0 : -r.x,
    originY: r.y === 0 ? 0 : -r.y,
  }
}

/** run line boxes of a match projected into slide px, each carrying the
 *  accumulated rotation of its ancestor-group chain plus the element's own —
 *  the overlay geometry the canvas draws through nested Konva containers.
 *  The walk starts UNflipped: hit.map (findNodeBox) already carries the top
 *  node's own flip as a scale(-1) matrix, which mirrors the rect positions
 *  and conjugates the segment rotations as one composed transform.
 *  Empty when the layout cannot be boxed (callers fall back to the element
 *  outline via findNodeBox + projectRect). */
export function matchStageRects(slide: RenderSlide | undefined, m: FindMatch): StageRect[] {
  const hit = findNodeBox(slide, m.sourceId)
  if (!hit) return []
  const segs: NodeSegment[] = []
  nodeSegments(hit.node, 0, 0, false, false, IDENTITY, segs)
  const rects: StageRect[] = []
  let off = 0
  for (const s of segs) {
    const sStart = off
    off += s.text.length
    if (!s.rect || !s.map || off <= m.start || sStart >= m.end) continue
    const a = Math.max(m.start, sStart) - sStart
    const b = Math.min(m.end, off) - sStart
    const len = s.text.length
    if (len <= 0) continue
    // partial-run coverage interpolates inside the run's box — exact at run
    // boundaries, approximate inside a variable-width run
    const f0 = a / len
    const f1 = b / len
    rects.push(
      projectRect(composeAffine(hit.map, s.map), {
        x: s.rect.x + f0 * s.rect.w,
        y: s.rect.y,
        w: (f1 - f0) * s.rect.w,
        h: s.rect.h,
      }),
    )
  }
  return rects
}

/** the element-outline fallback for a match, rotated by its full chain */
export function matchStageOutline(
  slide: RenderSlide | undefined,
  m: FindMatch,
): StageRect | undefined {
  const hit = findNodeBox(slide, m.sourceId)
  if (!hit) return undefined
  return projectRect(hit.map, { x: 0, y: 0, w: hit.node.box.w, h: hit.node.box.h })
}

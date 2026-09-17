/** Find logic (src/renderer/find-matches.ts): hit ranges under both conditions,
 *  deck-wide per-hit matching, and run-rect mapping for the highlight overlay. */
import { describe, it, expect } from 'vitest'
import type {
  GroupRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
  TableRenderNode,
} from '@airy-office/pptx-render'
import {
  buildMatches,
  findNodeBox,
  hitRanges,
  matchFocusBox,
  matchRects,
  matchStageOutline,
  matchStageRects,
  stepMatchIndex,
} from '../src/renderer/find-matches'

// ── fixtures: hand-built render trees (only the fields the walker reads) ──

const box = (x: number, y: number, w: number, h: number, extra: object = {}) => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
  ...extra,
})

function textNode(
  id: string,
  lines: Array<{ top: number; height: number; runs: Array<[string, number, number]> }>,
  extra: Record<string, unknown> = {},
): RenderNode {
  return {
    id,
    type: 'text',
    sourceId: id,
    box: box(100, 200, 400, 300),
    text: {
      lines: lines.map((l) => ({
        top: l.top,
        height: l.height,
        paraStart: true,
        runs: l.runs.map(([text, x, widthPx]) => ({
          text,
          x,
          baselineY: 0,
          widthPx,
          fontFamily: 'Arial',
          fontSizePx: 16,
          color: '000000',
          bold: false,
          italic: false,
          underline: false,
        })),
      })),
      insets: { l: 10, t: 6, r: 4, b: 4 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 40,
      wrap: true,
    },
    ...extra,
  } as unknown as RenderNode
}

const slideOf = (...nodes: RenderNode[]): RenderSlide =>
  ({ widthPx: 960, heightPx: 540, nodes }) as unknown as RenderSlide

describe('hitRanges', () => {
  const c = (matchCase: boolean, wholeWord: boolean) => ({ matchCase, wholeWord })

  it('finds every occurrence case-insensitively by default', () => {
    expect(hitRanges('Hello world, hello!', 'hello', c(false, false))).toEqual([
      [0, 5],
      [13, 18],
    ])
  })

  it('matchCase restricts to exact-case hits', () => {
    expect(hitRanges('Hello hello', 'hello', c(true, false))).toEqual([[6, 11]])
  })

  it('wholeWord rejects embedded occurrences (Unicode word chars)', () => {
    expect(hitRanges('cat cats catalog cat. cat-', 'cat', c(false, true))).toEqual([
      [0, 3],
      [17, 20],
      [22, 25],
    ])
  })

  it('wholeWord treats CJK neighbors as word characters', () => {
    expect(hitRanges('リンゴappleゴリラ', 'apple', c(false, true))).toEqual([])
    expect(hitRanges('リンゴ apple ゴリラ', 'apple', c(false, true))).toEqual([[4, 9]])
  })

  it('case-folding never shifts offsets (length-preserving)', () => {
    // 'İ' lowercases to a longer sequence; the folded haystack must keep offsets
    expect(hitRanges('İSTANBUL istanbul', 'istanbul', c(false, false))).toEqual([[9, 17]])
  })
})

describe('stepMatchIndex (find panel cursor math)', () => {
  it('enters at the first or last match without a cursor, then wraps', () => {
    expect(stepMatchIndex(-1, 1, 5)).toBe(0)
    expect(stepMatchIndex(-1, -1, 5)).toBe(4)
    expect(stepMatchIndex(0, 1, 5)).toBe(1)
    expect(stepMatchIndex(4, 1, 5)).toBe(0) // wrap forward
    expect(stepMatchIndex(0, -1, 5)).toBe(4) // wrap backward
    expect(stepMatchIndex(2, -1, 5)).toBe(1)
  })

  it('after a single replace consumed the cursor item, find-next lands on the following one', () => {
    // doReplace sets cursor = cursor - 1 after consuming the current match;
    // the next findWith(1) must return the index that followed the consumed
    // item, wrapping at the end
    const afterReplace = 2 - 1
    expect(stepMatchIndex(afterReplace, 1, 5)).toBe(2)
    expect(stepMatchIndex(0 - 1, 1, 5)).toBe(0)
    expect(stepMatchIndex(4 - 1, 1, 5)).toBe(4) // consumed the last item
  })

  it('returns -1 without matches', () => {
    expect(stepMatchIndex(-1, 1, 0)).toBe(-1)
    expect(stepMatchIndex(0, 1, 0)).toBe(-1)
  })
})

describe('buildMatches', () => {
  it('yields one match per hit with slide + element + range', () => {
    const slides = [
      slideOf(textNode('t1', [{ top: 0, height: 20, runs: [['cat cats', 0, 80]] }])),
      slideOf(textNode('t2', [{ top: 0, height: 20, runs: [['a cat', 0, 50]] }])),
    ]
    expect(buildMatches(slides, 'cat', { matchCase: false, wholeWord: true })).toEqual([
      { slideIndex: 0, sourceId: 't1', start: 0, end: 3 },
      { slideIndex: 1, sourceId: 't2', start: 2, end: 5 },
    ])
  })

  it('skips decoration nodes and walks group children', () => {
    const child = textNode('inner', [{ top: 0, height: 20, runs: [['target', 0, 60]] }])
    child.box = box(20, 30, 100, 40)
    const group: RenderNode = {
      id: 'g',
      type: 'group',
      sourceId: 'g',
      box: box(50, 60, 200, 100),
      children: [child],
    } as unknown as GroupRenderNode
    const deco = textNode('deco', [{ top: 0, height: 20, runs: [['target', 0, 60]] }])
    deco.decoration = true
    const slides = [slideOf(deco, group)]
    // matches inside a group are attributed to the group element (selection is
    // element-granularity), like the pre-threading behavior
    expect(buildMatches(slides, 'target', { matchCase: false, wholeWord: false })).toEqual([
      { slideIndex: 0, sourceId: 'g', start: 0, end: 6 },
    ])
    // the child itself still resolves with the accumulated offset
    expect(findNodeBox(slides[0], 'inner')).toMatchObject({ x: 70, y: 90 })
  })
})

describe('matchRects', () => {
  it('boxes the covered run (insets + run x, line box height)', () => {
    const n = textNode('t', [{ top: 12, height: 24, runs: [['hello world', 0, 110]] }])
    expect(matchRects(n, 0, 5)).toEqual([{ x: 10 + 0, y: 6 + 12, w: (5 / 11) * 110, h: 24 }])
  })

  it('spans two runs as two rects and interpolates partial coverage', () => {
    const n = textNode('t', [
      {
        top: 5,
        height: 20,
        runs: [
          ['foo ', 0, 40],
          ['bar!', 44, 40],
        ],
      },
    ])
    const rects = matchRects(n, 2, 6) // "o ba"
    expect(rects).toEqual([
      { x: 10 + (2 / 4) * 40, y: 6 + 5, w: (2 / 4) * 40, h: 20 },
      { x: 10 + 44, y: 6 + 5, w: (2 / 4) * 40, h: 20 },
    ])
  })

  it('maps table-cell hits through the cell origin', () => {
    const cell = {
      x: 30,
      y: 40,
      w: 100,
      h: 30,
      text: {
        lines: [
          {
            top: 4,
            height: 18,
            paraStart: true,
            runs: [{ text: 'hit', x: 2, widthPx: 30, baselineY: 0 }],
          },
        ],
        insets: { l: 5, t: 3, r: 2, b: 2 },
        fontScale: 1,
        contentHeight: 20,
        wrap: true,
      },
    }
    const table: RenderNode = {
      id: 'tbl',
      type: 'table',
      sourceId: 'tbl',
      box: box(10, 20, 300, 100),
      cells: [cell, { ...cell, x: 140, text: undefined }],
      gridX: [0, 100],
      gridY: [0, 30],
    } as unknown as TableRenderNode
    const rects = matchRects(table, 0, 3)
    expect(rects).toEqual([{ x: 30 + 5 + 2, y: 40 + 3 + 4, w: 30, h: 18 }])
  })

  it('vertical and warped layouts return no rects (element-box fallback)', () => {
    const vert = textNode('v', [{ top: 0, height: 20, runs: [['abc', 0, 30]] }], {})
    ;(vert as ShapeRenderNode).text!.vert = 'eaVert'
    expect(matchRects(vert, 0, 3)).toEqual([])
    const warped = textNode('w', [{ top: 0, height: 20, runs: [['abc', 0, 30]] }], {})
    ;(warped as ShapeRenderNode).text!.txWarp = { prst: 'textArchUp' }
    expect(matchRects(warped, 0, 3)).toEqual([])
  })
})

describe('matchFocusBox', () => {
  it('unions run rects into a slide-px box offset by the element origin', () => {
    const n = textNode('t', [
      {
        top: 10,
        height: 22,
        runs: [
          ['ab', 0, 20],
          ['cd', 20, 20],
        ],
      },
    ])
    const slide = slideOf(n)
    const m = { slideIndex: 0, sourceId: 't', start: 0, end: 4 }
    expect(matchFocusBox(slide, m)).toEqual({ x: 100 + 10, y: 200 + 6 + 10, w: 40, h: 22 })
  })

  it('falls back to the element box for unboxable layouts', () => {
    const n = textNode('t', [{ top: 0, height: 20, runs: [['abc', 0, 30]] }])
    ;(n as ShapeRenderNode).text!.vert = 'vert'
    const slide = slideOf(n)
    const m = { slideIndex: 0, sourceId: 't', start: 0, end: 3 }
    expect(matchFocusBox(slide, m)).toEqual({ x: 100, y: 200, w: 400, h: 300 })
  })
})

describe('matchStageRects (ancestor-group rotation chain)', () => {
  const cond = { matchCase: false, wholeWord: false }

  /** render a StageRect's top-left corner the way the CSS overlay does */
  function renderedCorner(r: {
    x: number
    y: number
    rotation: number
    originX: number
    originY: number
  }) {
    const ox = r.x + r.originX
    const oy = r.y + r.originY
    const rad = (r.rotation * Math.PI) / 180
    return {
      x: ox + (r.x - ox) * Math.cos(rad) - (r.y - oy) * Math.sin(rad),
      y: oy + (r.x - ox) * Math.sin(rad) + (r.y - oy) * Math.cos(rad),
    }
  }

  it('places text inside a group rotated 90° at the rotated position', () => {
    const child = textNode('inner', [{ top: 5, height: 20, runs: [['target', 0, 60]] }])
    child.box = box(20, 10, 40, 20)
    const group: RenderNode = {
      id: 'g',
      type: 'group',
      sourceId: 'g',
      box: box(200, 100, 80, 40, { rotationDeg: 90 }),
      children: [child],
    } as unknown as GroupRenderNode
    const slides = [slideOf(group)]
    const m = buildMatches(slides, 'target', cond)[0]!
    const rects = matchStageRects(slides[0], m)
    // run rect group-local: x = child origin(20) + inset l(10) + run x(0) = 30,
    // y = 10 + 6 + line top(5) = 21; the child is unrotated so its map is
    // identity and the chain is T(200,100) ∘ R90 about the group center
    // (40,20) → translation (260,80): the rect sits at (290,101) rotating
    // about (260,80), i.e. transform-origin (−30,−21) relative to the rect
    expect(rects).toEqual([
      { x: 290, y: 101, w: 60, h: 20, rotation: 90, originX: -30, originY: -21 },
    ])
    // rendering that rect lands its top-left corner at the affine-projected
    // spot: rotate (30,21) 90° about (40,20), then translate (200,100) → (239,110)
    expect(renderedCorner(rects[0]!)).toEqual({ x: expect.closeTo(239), y: expect.closeTo(110) })
  })

  it('accumulates rotations of nested groups (outer plain, inner 90°)', () => {
    const child = textNode('t', [{ top: 4, height: 18, runs: [['target', 0, 30]] }])
    child.box = box(5, 5, 30, 10)
    const inner: RenderNode = {
      id: 'in',
      type: 'group',
      sourceId: 'in',
      box: box(20, 10, 40, 20, { rotationDeg: 90 }),
      children: [child],
    } as unknown as GroupRenderNode
    const outer: RenderNode = {
      id: 'out',
      type: 'group',
      sourceId: 'out',
      box: box(200, 100, 80, 40),
      children: [inner],
    } as unknown as GroupRenderNode
    const slides = [slideOf(outer)]
    const m = buildMatches(slides, 'target', cond)[0]!
    // run rect outer-local: x = inner origin(20) + child origin(5) + inset(10)
    // = 35, y = 10 + 5 + 6 + 4 = 25; the inner group's rotation pivots its
    // center (40,20) in outer-local, the text child adds none:
    // chain T(200,100) ∘ R90 about (40,20) → t = (260,80)
    expect(matchStageRects(slides[0], m)).toEqual([
      { x: 295, y: 105, w: 30, h: 18, rotation: 90, originX: -35, originY: -25 },
    ])
  })

  it('keeps plain top-level hits unrotated at the element origin', () => {
    const n = textNode('t', [{ top: 12, height: 24, runs: [['hello world', 0, 110]] }])
    const slide = slideOf(n)
    const m = buildMatches([slide], 'hello', cond)[0]!
    expect(matchStageRects(slide, m)).toEqual([
      { x: 100 + 10, y: 200 + 6 + 12, w: 50, h: 24, rotation: 0, originX: 0, originY: 0 },
    ])
  })

  it('conjugates a child rotation under a flipped ancestor group (Konva parity)', () => {
    const child = textNode('inner', [{ top: 5, height: 20, runs: [['target', 0, 30]] }])
    child.box = box(20, 10, 40, 20, { rotationDeg: 30 })
    const group: RenderNode = {
      id: 'g',
      type: 'group',
      sourceId: 'g',
      box: box(200, 100, 80, 40, { flipH: true }),
      children: [child],
    } as unknown as GroupRenderNode
    const slides = [slideOf(group)]
    const m = buildMatches(slides, 'target', cond)[0]!

    // Konva-equivalent ground truth: each StaticNode Group applies
    // T(pos+center) ∘ R ∘ S ∘ T(−center) (boxPivotProps: position=center,
    // offset=center, rotation, scaleX/scaleY from the flip)
    const mul = (m1: number[], m2: number[]): number[] => [
      m1[0]! * m2[0]! + m1[2]! * m2[1]!,
      m1[1]! * m2[0]! + m1[3]! * m2[1]!,
      m1[0]! * m2[2]! + m1[2]! * m2[3]!,
      m1[1]! * m2[2]! + m1[3]! * m2[3]!,
      m1[0]! * m2[4]! + m1[2]! * m2[5]! + m1[4]!,
      m1[1]! * m2[4]! + m1[3]! * m2[5]! + m1[5]!,
    ]
    const konvaNodeMap = (b: {
      x: number
      y: number
      w: number
      h: number
      rotationDeg: number
      flipH?: boolean
    }): number[] => {
      const rad = (b.rotationDeg * Math.PI) / 180
      const cos = Math.cos(rad)
      const sin = Math.sin(rad)
      const cx = b.w / 2
      const cy = b.h / 2
      // linear part R ∘ S with S = diag(flipH ? -1 : 1, 1)
      const rotateScale = [cos * (b.flipH ? -1 : 1), sin * (b.flipH ? -1 : 1), -sin, cos, 0, 0]
      const translate = [1, 0, 0, 1, b.x + cx, b.y + cy]
      const unOffset = [1, 0, 0, 1, -cx, -cy]
      return mul(mul(translate, rotateScale), unOffset)
    }
    const total = mul(konvaNodeMap(group.box as never), konvaNodeMap(child.box as never))
    // run rect in child-local px: inset l(10) + run x(0), inset t(6) + top(5)
    const rect = { x: 10, y: 11, w: 30, h: 20 }
    const expected = [
      { x: rect.x, y: rect.y },
      { x: rect.x + rect.w, y: rect.y },
      { x: rect.x, y: rect.y + rect.h },
      { x: rect.x + rect.w, y: rect.y + rect.h },
    ].map((p) => ({
      x: total[0]! * p.x + total[2]! * p.y + total[4]!,
      y: total[1]! * p.x + total[3]! * p.y + total[5]!,
    }))

    const rects = matchStageRects(slides[0], m)
    expect(rects.length).toBe(1)
    // render the StageRect the way the CSS overlay does and compare corner SETS
    const stage = rects[0]!
    const ox = stage.x + stage.originX
    const oy = stage.y + stage.originY
    const rad = (stage.rotation * Math.PI) / 180
    const cos = Math.cos(rad)
    const sin = Math.sin(rad)
    const rendered = [
      { x: stage.x, y: stage.y },
      { x: stage.x + stage.w, y: stage.y },
      { x: stage.x, y: stage.y + stage.h },
      { x: stage.x + stage.w, y: stage.y + stage.h },
    ].map((p) => ({
      x: ox + (p.x - ox) * cos - (p.y - oy) * sin,
      y: oy + (p.x - ox) * sin + (p.y - oy) * cos,
    }))
    const key = (p: { x: number; y: number }) => `${p.x.toFixed(6)},${p.y.toFixed(6)}`
    expect(new Set(rendered.map(key))).toEqual(new Set(expected.map(key)))
  })

  it('conjugates through a rotated+flipped group containing a rotated child', () => {
    const child = textNode('inner', [{ top: 4, height: 18, runs: [['target', 0, 25]] }])
    child.box = box(10, 6, 30, 16, { rotationDeg: -25 })
    const group: RenderNode = {
      id: 'g',
      type: 'group',
      sourceId: 'g',
      box: box(300, 150, 90, 50, { rotationDeg: 40, flipH: true }),
      children: [child],
    } as unknown as GroupRenderNode
    const slides = [slideOf(group)]
    const m = buildMatches(slides, 'target', cond)[0]!
    const rects = matchStageRects(slides[0], m)
    expect(rects.length).toBe(1)
    // a mirrored ancestor reverses the child's rotation direction: the
    // chain R40 ∘ mirror ∘ R(−25) equals R65 ∘ mirror, so the projected
    // highlight rect rotates by 65°
    expect(Math.abs(Math.abs(rects[0]!.rotation) - 65)).toBeLessThan(1e-6)
  })

  it('outlines a rotated element for unboxable layouts (vertical text)', () => {
    const n = textNode('v', [{ top: 0, height: 20, runs: [['abc', 0, 30]] }])
    n.box = box(100, 200, 400, 300, { rotationDeg: 90 })
    ;(n as ShapeRenderNode).text!.vert = 'eaVert'
    const slide = slideOf(n)
    const m = buildMatches([slide], 'abc', cond)[0]!
    expect(matchStageRects(slide, m)).toEqual([])
    // chain T(100,200) ∘ R90 about (200,150): translation = (100,200)+(200,150)−R(200,150)
    // R90(200,150) = (−150,200) → t = (450,150); outline at (0+450, 0+150)
    expect(matchStageOutline(slide, m)).toEqual({
      x: 450,
      y: 150,
      w: 400,
      h: 300,
      rotation: 90,
      originX: 0,
      originY: 0,
    })
  })
})

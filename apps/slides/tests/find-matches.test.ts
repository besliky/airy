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

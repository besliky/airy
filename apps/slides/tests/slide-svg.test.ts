/** Vector slide painting (src/renderer/slide-svg.ts): exported pages carry
 *  real <text>, geometry from the render tree, and the documented
 *  approximations stay out of the structure. */
import { describe, it, expect } from 'vitest'
import type {
  ChartRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
  TableRenderNode,
} from '@airy-office/pptx-render'
import { renderSlideSvg } from '../src/renderer/slide-svg'

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

const run = (text: string, x: number, baselineY: number, extra: object = {}) => ({
  text,
  x,
  baselineY,
  widthPx: text.length * 10,
  fontFamily: 'Carlito',
  fontSizePx: 18,
  color: '102030',
  bold: false,
  italic: false,
  underline: false,
  ...extra,
})

const slideOf = (...nodes: RenderNode[]): RenderSlide =>
  ({
    widthPx: 960,
    heightPx: 540,
    scale: 1,
    background: { kind: 'solid', color: 'FFFFFF' },
    nodes,
  }) as unknown as RenderSlide

describe('renderSlideSvg', () => {
  it('emits real <text> elements with font, color and decorations', () => {
    const node: RenderNode = {
      id: 't1',
      type: 'text',
      sourceId: 't1',
      box: box(100, 50, 400, 60),
      text: {
        lines: [{ top: 4, height: 24, paraStart: true, runs: [run('Quarterly Review', 8, 22)] }],
        insets: { l: 10, t: 6, r: 4, b: 4 },
        anchor: 'top',
        fontScale: 1,
        contentHeight: 30,
        wrap: true,
      },
    } as unknown as ShapeRenderNode
    const svg = renderSlideSvg(slideOf(node), new Map())
    expect(svg).toMatch(/^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" viewBox="0 0 960 540"/)
    // run coords are node-local (the group carries the box translation): insets + run offsets
    expect(svg).toContain(
      '<text x="18.00" y="28.00" font-family="Carlito" font-size="18.00" fill="#102030" xml:space="preserve">Quarterly Review</text>',
    )
    expect(svg).toContain('<g transform="translate(100.00 50.00)">')
  })

  it('escapes markup in text content', () => {
    const node: RenderNode = {
      id: 't2',
      type: 'text',
      sourceId: 't2',
      box: box(0, 0, 200, 40),
      text: {
        lines: [{ top: 0, height: 20, paraStart: true, runs: [run('a < b & "c"', 0, 16)] }],
        insets: { l: 0, t: 0, r: 0, b: 0 },
        anchor: 'top',
        fontScale: 1,
        contentHeight: 20,
        wrap: true,
      },
    } as unknown as ShapeRenderNode
    expect(renderSlideSvg(slideOf(node), new Map())).toContain('a &lt; b &amp; "c"')
  })

  it('paints shape geometry from pathData with fills, strokes and run highlights', () => {
    const node: RenderNode = {
      id: 's1',
      type: 'shape',
      sourceId: 's1',
      box: box(10, 20, 300, 120),
      presetGeometry: 'custom',
      pathData: 'M0 0L300 0L300 120L0 120Z',
      fill: { kind: 'solid', color: '112233' },
      stroke: { color: '445566', widthPx: 2, widthPt: 1.5 },
      text: {
        lines: [
          {
            top: 10,
            height: 26,
            paraStart: true,
            runs: [run('highlighted', 12, 30, { highlight: 'FFFF00' })],
          },
        ],
        insets: { l: 6, t: 4, r: 6, b: 4 },
        anchor: 'top',
        fontScale: 1,
        contentHeight: 40,
        wrap: true,
      },
    } as unknown as ShapeRenderNode
    const svg = renderSlideSvg(slideOf(node), new Map())
    expect(svg).toContain(
      '<path d="M0 0L300 0L300 120L0 120Z" fill="#112233" stroke="#445566" stroke-width="2.00"',
    )
    expect(svg).toContain('fill="#FFFF00"')
    expect(svg).toContain('>highlighted</text>')
  })

  it('gradient and image fills land in <defs> and reference them', () => {
    const gradient: RenderNode = {
      id: 'g1',
      type: 'shape',
      sourceId: 'g1',
      box: box(0, 0, 100, 100),
      presetGeometry: 'rect',
      fill: {
        kind: 'gradient',
        stops: [
          { pos: 0, color: 'FF0000' },
          { pos: 1, color: '0000FF' },
        ],
        angleDeg: 90,
      },
    } as unknown as ShapeRenderNode
    const svg = renderSlideSvg(slideOf(gradient), new Map())
    expect(svg).toContain('<linearGradient id="grad0"')
    expect(svg).toContain('fill="url(#grad0)"')
    expect(svg).toContain('stop-color="#FF0000"')
  })

  it('pictures embed the image and crop via srcRect; tables draw cells + text', () => {
    const pic: RenderNode = {
      id: 'p1',
      type: 'picture',
      sourceId: 'p1',
      box: box(0, 0, 200, 100),
      dataUrl: 'data:image/png;base64,AAAA',
      srcRect: { l: 0.25, t: 0, r: 0, b: 0 },
    } as unknown as PictureRenderNode
    const table: RenderNode = {
      id: 'tbl',
      type: 'table',
      sourceId: 'tbl',
      box: box(0, 150, 400, 80),
      gridX: [0, 200, 400],
      gridY: [0, 80],
      cells: [
        {
          x: 0,
          y: 0,
          w: 200,
          h: 80,
          row: 0,
          col: 0,
          merged: false,
          fill: { kind: 'solid', color: 'EEEEEE' },
          borders: { t: { color: '000000', widthPx: 1, widthPt: 0.75 } },
          text: {
            lines: [{ top: 8, height: 20, paraStart: true, runs: [run('cell', 4, 24)] }],
            insets: { l: 4, t: 4, r: 4, b: 4 },
            fontScale: 1,
            contentHeight: 30,
            wrap: true,
          },
        },
      ],
    } as unknown as TableRenderNode
    const svg = renderSlideSvg(slideOf(pic, table), new Map())
    // 25% left crop: visible 75% of the source fills the 200px frame → 266.67 wide, x=-66.67
    expect(svg).toContain('x="-66.67"')
    expect(svg).toContain('data:image/png;base64,AAAA')
    expect(svg).toContain('>cell</text>')
    expect(svg).toContain('stroke="#000000" stroke-width="1.00"')
  })

  it('charts draw primitives and their labels as text', () => {
    const chart: RenderNode = {
      id: 'c1',
      type: 'chart',
      sourceId: 'c1',
      box: box(0, 0, 300, 200),
      gridLines: [],
      axisLines: [],
      labels: [{ text: 'Q1', x: 10, y: 20, fontSizePx: 12, color: '333333' }],
      bars: [{ x: 10, y: 100, w: 40, h: 60, color: '4472C4' }],
      polylines: [],
      markers: [],
      swatches: [],
      wedges: [
        { cx: 150, cy: 100, outerR: 60, innerR: 0, startDeg: 0, sweepDeg: 120, color: 'ED7D31' },
      ],
    } as unknown as ChartRenderNode
    const svg = renderSlideSvg(slideOf(chart), new Map())
    expect(svg).toContain('fill="#4472C4"')
    expect(svg).toContain('A 60.00 60.00 0 0 1')
    expect(svg).toContain('>Q1</text>')
  })

  it('rotated and flipped nodes transform their geometry but keep text upright', () => {
    const node: RenderNode = {
      id: 'f1',
      type: 'shape',
      sourceId: 'f1',
      box: box(50, 60, 200, 80, { rotationDeg: 15, flipH: true }),
      presetGeometry: 'rect',
      fill: { kind: 'solid', color: '001122' },
      text: {
        lines: [{ top: 2, height: 20, paraStart: true, runs: [run('mirror', 5, 18)] }],
        insets: { l: 0, t: 0, r: 0, b: 0 },
        anchor: 'top',
        fontScale: 1,
        contentHeight: 24,
        wrap: true,
      },
    } as unknown as ShapeRenderNode
    const svg = renderSlideSvg(slideOf(node), new Map())
    expect(svg).toContain('rotate(15 100.00 40.00)')
    expect(svg).toContain('translate(200.00 0) scale(-1 1)')
    // flipped text mirrors its anchor inside the box: 200 - (5 + 6*10) = 135
    expect(svg).toContain('x="135.00"')
  })

  it('mirrors flipped pictures, tables, and whole groups about the box center', () => {
    const pic: RenderNode = {
      id: 'p2',
      type: 'picture',
      sourceId: 'p2',
      box: box(0, 0, 200, 100, { flipH: true }),
      dataUrl: 'data:image/png;base64,BBBB',
    } as unknown as PictureRenderNode
    const table: RenderNode = {
      id: 'tbl2',
      type: 'table',
      sourceId: 'tbl2',
      box: box(0, 150, 400, 80, { flipV: true }),
      gridX: [0, 400],
      gridY: [0, 80],
      cells: [
        {
          x: 0,
          y: 0,
          w: 400,
          h: 80,
          row: 0,
          col: 0,
          merged: false,
          fill: { kind: 'solid', color: 'EEEEEE' },
          text: {
            lines: [],
            insets: { l: 4, t: 4, r: 4, b: 4 },
            fontScale: 1,
            contentHeight: 0,
            wrap: true,
          },
        },
      ],
    } as unknown as TableRenderNode
    const group: RenderNode = {
      id: 'grp',
      type: 'group',
      sourceId: 'grp',
      box: box(10, 20, 100, 50, { flipH: true }),
      children: [
        {
          id: 'inner',
          type: 'shape',
          sourceId: 'inner',
          box: box(0, 0, 100, 50),
          presetGeometry: 'rect',
          fill: { kind: 'solid', color: '445566' },
        } as unknown as ShapeRenderNode,
      ],
    } as unknown as RenderNode
    const svg = renderSlideSvg(slideOf(pic, table, group), new Map())
    // picture: mirror about the box center (w=200), image inside the flip group
    expect(svg).toContain('<g transform="translate(200.00 0) scale(-1 1)"><g><image ')
    // table: flipV mirrors about the box center height (h=80)
    expect(svg).toContain('<g transform="translate(0 80.00) scale(1 -1)"><rect ')
    // group: the whole child subtree flips with the group
    expect(svg).toContain(
      '<g transform="translate(100.00 0) scale(-1 1)"><g transform="translate(0.00 0.00)">',
    )
  })
  it('draws a 360° pie wedge as a closed full circle', () => {
    const chart: RenderNode = {
      id: 'c2',
      type: 'chart',
      sourceId: 'c2',
      box: box(0, 0, 300, 200),
      gridLines: [],
      axisLines: [],
      labels: [],
      bars: [],
      polylines: [],
      markers: [],
      swatches: [],
      wedges: [
        { cx: 150, cy: 100, outerR: 60, innerR: 0, startDeg: 0, sweepDeg: 360, color: 'ED7D31' },
      ],
    } as unknown as ChartRenderNode
    const svg = renderSlideSvg(slideOf(chart), new Map())
    // start == end would draw nothing; the full circle is two half arcs
    expect(svg.match(/A 60\.00 60\.00 0 1 1/g)).toHaveLength(2)
    expect(svg).toContain(
      'M 150.00 40.00 A 60.00 60.00 0 1 1 150.00 160.00 A 60.00 60.00 0 1 1 150.00 40.00 Z',
    )
  })
})

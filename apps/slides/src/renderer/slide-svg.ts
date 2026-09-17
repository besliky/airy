/**
 * Vector slide painting for PDF export: renders a RenderSlide to a
 * self-contained inline <svg> whose text is real <text> elements, so
 * printToPDF output has selectable, searchable text instead of page bitmaps.
 *
 * Geometry comes from the same render tree the Konva canvas draws (pathData is
 * SVG d-strings, chart draw primitives, per-run text layout), positioned with
 * the same box/flip/rotation rules the canvas uses (rotation and flips pivot
 * on the box center; a flipped shape's text stays readable, mirroring
 * NodeBody's counter-flip).
 *
 * Deliberately approximated (visual extras, not structure): WordArt
 * warps/extrusion, run shadows/glows/reflections, picture pixel filters
 * (duotone/clrChange/lum render the unfiltered image), pattern fills
 * (diagonal-stripe approximation), smoothed chart polylines export as their
 * control points, and rotated-text anchors follow the box center. When SVG
 * assembly throws, the exporter falls back to the raster page for that slide.
 */
import type {
  ChartRenderNode,
  GroupRenderNode,
  PictureRenderNode,
  RenderFill,
  RenderNode,
  RenderSlide,
  RenderStroke,
  ShapeRenderNode,
  TableRenderNode,
  TextLine,
} from '@airy-office/pptx-render'

let defSeq = 0

const escText = (s: string) =>
  s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!)
const escAttr = (s: string) => escText(s).replace(/"/g, '&quot;')

/** flat "x,y x,y" attribute value from a [x0,y0,x1,y1,…] array */
const ptsAttr = (points: number[]): string => {
  const out: string[] = []
  for (let i = 0; i + 1 < points.length; i += 2) {
    out.push(`${points[i]!.toFixed(2)},${points[i + 1]!.toFixed(2)}`)
  }
  return out.join(' ')
}

/** #RRGGBB / RRGGBB / RRGGBBAA / rgba() → CSS color ('' for none) */
function cssColor(c: string | undefined): string {
  if (!c || c === 'none') return ''
  if (c.startsWith('#') || c.startsWith('rgb')) return c
  if (/^[0-9A-Fa-f]{8}$/.test(c)) {
    const a = parseInt(c.slice(6, 8), 16) / 255
    return `rgba(${parseInt(c.slice(0, 2), 16)},${parseInt(c.slice(2, 4), 16)},${parseInt(
      c.slice(4, 6),
      16,
    )},${a.toFixed(3)})`
  }
  return `#${c}`
}

/** gradient angle → unit-square endpoints (0° = left→right, 90° = top→bottom) */
function gradientVector(angleDeg: number): { x1: number; y1: number; x2: number; y2: number } {
  const rad = (angleDeg * Math.PI) / 180
  const dx = Math.cos(rad)
  const dy = Math.sin(rad)
  // ramp length = the unit square's projection onto the direction
  const len = Math.abs(dx) + Math.abs(dy)
  const hx = (dx * len) / 2
  const hy = (dy * len) / 2
  return { x1: 0.5 - hx, y1: 0.5 - hy, x2: 0.5 + hx, y2: 0.5 + hy }
}

/** resolved paint for one surface */
interface FillResult {
  /** SVG fill value ('none' when a solid 'none' or when an image underlay carries the pixels) */
  fill: string
  fillOpacity: number
  /** image fills draw as clipped <image> markup placed under the geometry */
  underlay?: string
}

/**
 * Resolve a fill. Gradient/pattern paints append their <defs> entries to
 * `defs`; image fills come back as `underlay` markup (a clipped <image>).
 */
function paintFill(
  fill: RenderFill | undefined,
  boxW: number,
  boxH: number,
  defs: string[],
  alpha = 1,
): FillResult {
  if (!fill || fill.kind === 'none') return { fill: 'none', fillOpacity: 1 }
  if (fill.kind === 'solid') {
    return { fill: cssColor(fill.color) || 'none', fillOpacity: alpha }
  }
  if (fill.kind === 'gradient') {
    const id = `grad${defSeq++}`
    const stops = fill.stops
      .map((s) => `<stop offset="${(s.pos * 100).toFixed(2)}%" stop-color="${cssColor(s.color)}"/>`)
      .join('')
    if (fill.radial) {
      const cx = ((fill.center?.x ?? 0.5) * 100).toFixed(1)
      const cy = ((fill.center?.y ?? 0.5) * 100).toFixed(1)
      defs.push(
        `<radialGradient id="${id}" cx="${cx}%" cy="${cy}%" r="75%">${stops}</radialGradient>`,
      )
    } else {
      const v = gradientVector(fill.angleDeg)
      defs.push(
        `<linearGradient id="${id}" x1="${v.x1.toFixed(4)}" y1="${v.y1.toFixed(4)}" x2="${v.x2.toFixed(4)}" y2="${v.y2.toFixed(4)}">${stops}</linearGradient>`,
      )
    }
    return { fill: `url(#${id})`, fillOpacity: alpha }
  }
  if (fill.kind === 'image' && fill.dataUrl) {
    if (fill.mode === 'tile' && fill.tile) {
      const id = `tile${defSeq++}`
      const tw = Math.max(fill.tile.scaleX, 1)
      const th = Math.max(fill.tile.scaleY, 1)
      defs.push(
        `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${tw.toFixed(2)}" height="${th.toFixed(2)}">` +
          `<image href="${escAttr(fill.dataUrl)}" width="${tw.toFixed(2)}" height="${th.toFixed(2)}"/></pattern>`,
      )
      return { fill: `url(#${id})`, fillOpacity: fill.alpha ?? alpha }
    }
    // stretched: the image maps into its fillRect sub-rect of the box
    const l = fill.fillRect?.l ?? 0
    const t = fill.fillRect?.t ?? 0
    const r = fill.fillRect?.r ?? 0
    const b = fill.fillRect?.b ?? 0
    const id = `clip${defSeq++}`
    defs.push(
      `<clipPath id="${id}"><rect x="0" y="0" width="${boxW.toFixed(2)}" height="${boxH.toFixed(2)}"/></clipPath>`,
    )
    const x = l < 0 ? l * boxW : 0
    const y = t < 0 ? t * boxH : 0
    const w = (1 + r - l) * boxW
    const h = (1 + b - t) * boxH
    const opacity = fill.alpha !== undefined && fill.alpha < 1 ? ` opacity="${fill.alpha}"` : ''
    return {
      fill: 'none',
      fillOpacity: 1,
      underlay:
        `<g clip-path="url(#${id})">` +
        `<image href="${escAttr(fill.dataUrl)}" x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" preserveAspectRatio="none"${opacity}/></g>`,
    }
  }
  if (fill.kind === 'pattern') {
    const id = `pat${defSeq++}`
    const cell = Math.max(fill.cellPx, 2)
    const fg = cssColor(fill.fg) || '#000'
    const bg = cssColor(fill.bg) || '#fff'
    defs.push(
      `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${cell}" height="${cell}">` +
        `<rect width="${cell}" height="${cell}" fill="${bg}"/>` +
        `<rect width="${(cell / 2).toFixed(2)}" height="${(cell / 2).toFixed(2)}" fill="${fg}"/></pattern>`,
    )
    return { fill: `url(#${id})`, fillOpacity: alpha }
  }
  return { fill: 'none', fillOpacity: 1 }
}

/** stroke → SVG attributes */
function strokeAttrs(s: RenderStroke | undefined): string {
  if (!s) return ''
  const color = cssColor(s.color)
  if (!color) return ''
  const dash = s.dash?.length ? ` stroke-dasharray="${s.dash.join(' ')}"` : ''
  const cap = s.cap === 'round' ? 'round' : 'butt'
  const join = s.join === 'bevel' ? 'bevel' : s.join === 'miter' ? 'miter' : 'round'
  return ` stroke="${color}" stroke-width="${Math.max(s.widthPx, 0.1).toFixed(2)}"${dash} stroke-linecap="${cap}" stroke-linejoin="${join}"`
}

/** font attributes shared by slide runs and chart labels */
function fontAttrs(o: {
  fontFamily: string
  fontSizePx: number
  bold?: boolean
  italic?: boolean
  letterSpacingPx?: number
}): string {
  return (
    `font-family="${escAttr(o.fontFamily)}" font-size="${o.fontSizePx.toFixed(2)}"` +
    (o.bold ? ' font-weight="bold"' : '') +
    (o.italic ? ' font-style="italic"' : '') +
    ((o.letterSpacingPx ?? 0) ? ` letter-spacing="${(o.letterSpacingPx ?? 0).toFixed(2)}"` : '')
  )
}

/** one laid-out line → positioned <text> elements (these become real PDF text) */
function lineText(
  line: TextLine,
  ox: number,
  oy: number,
  vert: boolean,
  flipW: number,
  flipH: number,
): string {
  let out = ''
  for (const run of line.runs) {
    if (!run.text) continue
    const left = ox + run.x
    const top = oy + line.top
    const ascentFromTop = run.baselineY - line.top
    // flipped container: mirror the run's anchor inside the box (glyphs stay readable)
    const rx = flipW > 0 ? flipW - (left + run.widthPx) : left
    const ry = flipH > 0 ? flipH - (top + line.height) : top
    const rby = flipH > 0 ? ry + (line.height - ascentFromTop) : oy + run.baselineY
    const deco = [run.underline ? 'underline' : '', run.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' ')
    // PowerPoint paints run highlights over the whole line box
    const hl = run.highlight
      ? `<rect x="${rx.toFixed(2)}" y="${ry.toFixed(2)}" width="${run.widthPx.toFixed(2)}" height="${line.height.toFixed(2)}" fill="${cssColor(run.highlight)}"/>`
      : ''
    const rotate = run.rotate90
      ? ` transform="rotate(90 ${rx.toFixed(2)} ${rby.toFixed(2)})"`
      : run.rotate270
        ? ` transform="rotate(-90 ${rx.toFixed(2)} ${rby.toFixed(2)})"`
        : ''
    out +=
      hl +
      `<text x="${rx.toFixed(2)}" y="${rby.toFixed(2)}" ${fontAttrs({
        fontFamily: run.fontFamily,
        fontSizePx: run.fontSizePx,
        bold: run.bold,
        italic: run.italic,
        letterSpacingPx: (run.letterSpacingPx ?? 0) + (run.justifyExtraPx ?? 0),
      })}${deco ? ` text-decoration="${deco}"` : ''} fill="${cssColor(run.color) || '#000'}"${vert ? ' writing-mode="tb"' : ''}${rotate} xml:space="preserve">${escText(run.text)}</text>`
  }
  return out
}

/** a text layout (shape body / table cell) → highlight + <text> elements */
function layoutText(
  text: ShapeRenderNode['text'],
  ox: number,
  oy: number,
  flipW = 0,
  flipH = 0,
): string {
  if (!text) return ''
  const vert = !!text.vert
  const il = text.insets?.l ?? 0
  const it = text.insets?.t ?? 0
  return text.lines.map((l) => lineText(l, ox + il, oy + it, vert, flipW, flipH)).join('')
}

/** arrowhead markup (triangle-style approximation of the OOXML head types) */
function arrowHead(
  x: number,
  y: number,
  angleRad: number,
  end: { type: string; widthPx: number; lengthPx: number },
): string {
  const w = Math.max(end.widthPx, 1)
  const l = Math.max(end.lengthPx, 1)
  const ux = Math.cos(angleRad)
  const uy = Math.sin(angleRad)
  const px = -uy
  const py = ux
  const tipX = x + ux * l
  const tipY = y + uy * l
  const pts: Array<[number, number]> =
    end.type === 'diamond'
      ? [
          [x, y],
          [(x + tipX) / 2 + px * w, (y + tipY) / 2 + py * w],
          [tipX, tipY],
          [(x + tipX) / 2 - px * w, (y + tipY) / 2 - py * w],
        ]
      : [
          [x + px * w, y + py * w],
          [tipX, tipY],
          [x - px * w, y - py * w],
        ]
  return `<polygon points="${ptsAttr(pts.flat())}"/>`
}

/** shape geometry + fills + stroke → SVG body in node-local coordinates */
function shapeGeometry(shape: ShapeRenderNode, defs: string[]): string {
  const { box } = shape
  const paint = paintFill(shape.fill, box.w, box.h, defs)
  const stroke = strokeAttrs(shape.stroke)
  const fillOpacity = paint.fillOpacity < 1 ? ` fill-opacity="${paint.fillOpacity}"` : ''
  const common = `fill="${paint.fill}"${fillOpacity}${stroke}`

  if (shape.line) {
    // connector: polyline with points in node-local space (flips pre-baked)
    const pts = shape.line.points
    const seg = (i: number, j: number) =>
      Math.atan2(pts[j * 2 + 1]! - pts[i * 2 + 1]!, pts[j * 2]! - pts[i * 2]!)
    const line = `<polyline points="${ptsAttr(pts)}" fill="none"${stroke}/>`
    const headFill = ` fill="${cssColor(shape.stroke?.color) || '#000'}" stroke="none"`
    const heads = shape.line.headEnd
      ? `<g${headFill}>${arrowHead(pts[0] ?? 0, pts[1] ?? 0, seg(0, 1), shape.line.headEnd)}</g>`
      : ''
    const tails = shape.line.tailEnd
      ? `<g${headFill}>${arrowHead(
          pts[pts.length - 2] ?? 0,
          pts[pts.length - 1] ?? 0,
          seg(pts.length / 2 - 2, pts.length / 2 - 1),
          shape.line.tailEnd,
        )}</g>`
      : ''
    return `${line}${heads}${tails}`
  }

  let geo = ''
  if (shape.pathData) geo += `<path d="${shape.pathData}" ${common}/>`
  if (shape.fillPathData)
    geo += `<path d="${shape.fillPathData}" fill="${paint.fill}"${fillOpacity} stroke="none"/>`
  if (shape.strokePathData) geo += `<path d="${shape.strokePathData}" fill="none"${stroke}/>`
  if (!geo && shape.polygonPoints) {
    geo += `<polygon points="${ptsAttr(shape.polygonPoints)}" ${common}/>`
  }
  if (!geo) {
    const preset = shape.presetGeometry ?? ''
    if (preset === 'ellipse' || preset === 'oval') {
      geo = `<ellipse cx="${(box.w / 2).toFixed(2)}" cy="${(box.h / 2).toFixed(2)}" rx="${(box.w / 2).toFixed(2)}" ry="${(box.h / 2).toFixed(2)}" ${common}/>`
    } else if (shape.cornerRadiusPx) {
      geo = `<rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" rx="${shape.cornerRadiusPx.toFixed(2)}" ${common}/>`
    } else {
      geo = `<rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" ${common}/>`
    }
  }
  return `${paint.underlay ?? ''}${geo}`
}

/** chart primitives → SVG body in node-local coordinates */
function chartGeometry(chart: ChartRenderNode, defs: string[]): string {
  const { box } = chart
  const bg = paintFill(chart.bgFill, box.w, box.h, defs)
  let out = `<rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" fill="${bg.fill}"/>`
  if (chart.plotRect) {
    const p = chart.plotRect
    const pf = paintFill(p.fill, p.w, p.h, defs)
    out += `<rect x="${p.x.toFixed(2)}" y="${p.y.toFixed(2)}" width="${p.w.toFixed(2)}" height="${p.h.toFixed(2)}" fill="${pf.fill}"${
      p.borderColor
        ? ` stroke="${cssColor(p.borderColor)}" stroke-width="${p.borderWidthPx ?? 1}"`
        : ''
    }/>`
  }
  out += chart.gridLines
    .map(
      (g) =>
        `<line x1="${g.x1.toFixed(2)}" y1="${g.y1.toFixed(2)}" x2="${g.x2.toFixed(2)}" y2="${g.y2.toFixed(2)}" stroke="${cssColor(g.color)}" stroke-width="${(g.widthPx ?? 1).toFixed(2)}"${g.dash?.length ? ` stroke-dasharray="${g.dash.join(' ')}"` : ''}/>`,
    )
    .join('')
  out += chart.axisLines
    .map(
      (a) =>
        `<line x1="${a.x1.toFixed(2)}" y1="${a.y1.toFixed(2)}" x2="${a.x2.toFixed(2)}" y2="${a.y2.toFixed(2)}" stroke="${cssColor(a.color)}" stroke-width="${a.widthPx.toFixed(2)}"/>`,
    )
    .join('')
  out += chart.bars
    .map(
      (b) =>
        `<rect x="${b.x.toFixed(2)}" y="${b.y.toFixed(2)}" width="${Math.max(b.w, 0).toFixed(2)}" height="${Math.max(b.h, 0).toFixed(2)}" fill="${cssColor(b.color)}"/>`,
    )
    .join('')
  out += (chart.wedges ?? [])
    .map(
      (w) =>
        `<path d="${wedgePath(w)}" fill="${w.noFill ? 'none' : cssColor(w.color)}"${
          w.stroke
            ? ` stroke="${cssColor(w.stroke)}" stroke-width="${w.strokeWidthPx ?? 1}"`
            : ' stroke="#ffffff" stroke-width="1"'
        }/>`,
    )
    .join('')
  out += (chart.paths ?? [])
    .map(
      (p) =>
        `<path d="${p.d}" fill="${cssColor(p.fill)}"${p.stroke ? ` stroke="${cssColor(p.stroke)}" stroke-width="${p.strokeWidthPx ?? 1}"` : ''}${p.dy ? ` transform="translate(0 ${p.dy.toFixed(2)})"` : ''}/>`,
    )
    .join('')
  out += chart.polylines
    .map((p) => {
      const shape = p.closed ? 'polygon' : 'polyline'
      const points = ptsAttr(p.points)
      return `<${shape} points="${points}" fill="${p.fill ? cssColor(p.fill) : 'none'}" stroke="${cssColor(p.color)}" stroke-width="${p.widthPx.toFixed(2)}"${p.dash?.length ? ` stroke-dasharray="${p.dash.join(' ')}"` : ''}/>`
    })
    .join('')
  out += chart.markers
    .map(
      (m) =>
        `<circle cx="${m.x.toFixed(2)}" cy="${m.y.toFixed(2)}" r="${m.r.toFixed(2)}" fill="${cssColor(m.color)}"/>`,
    )
    .join('')
  out += chart.swatches
    .map(
      (s) =>
        `<rect x="${s.x.toFixed(2)}" y="${s.y.toFixed(2)}" width="${s.w.toFixed(2)}" height="${s.h.toFixed(2)}" fill="${cssColor(s.color)}"/>`,
    )
    .join('')
  // labels become real PDF text too
  out += chart.labels
    .map((l) => {
      const y = l.y + l.fontSizePx
      const rot = l.rotationDeg
        ? ` transform="rotate(${l.rotationDeg} ${l.x.toFixed(2)} ${y.toFixed(2)})"`
        : ''
      return `<text x="${l.x.toFixed(2)}" y="${y.toFixed(2)}" ${fontAttrs({
        fontFamily: "-apple-system, 'Segoe UI', sans-serif",
        fontSizePx: l.fontSizePx,
        bold: l.bold,
        italic: l.italic,
      })} fill="${cssColor(l.color) || '#000'}"${rot}>${escText(l.text)}</text>`
    })
    .join('')
  if (chart.border) {
    out += `<rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" fill="none" stroke="${cssColor(chart.border.color)}" stroke-width="${chart.border.widthPx.toFixed(2)}"/>`
  }
  return out
}

/** pie/doughnut wedge → SVG arc path */
function wedgePath(w: NonNullable<ChartRenderNode['wedges']>[number]): string {
  const { cx, cy, outerR, innerR, startDeg, sweepDeg } = w
  const large = sweepDeg > 180 ? 1 : 0
  const a0 = ((startDeg - 90) * Math.PI) / 180
  const a1 = ((startDeg + sweepDeg - 90) * Math.PI) / 180
  const x0 = cx + outerR * Math.cos(a0)
  const y0 = cy + outerR * Math.sin(a0)
  const x1 = cx + outerR * Math.cos(a1)
  const y1 = cy + outerR * Math.sin(a1)
  if (innerR <= 0) {
    return `M ${cx.toFixed(2)} ${cy.toFixed(2)} L ${x0.toFixed(2)} ${y0.toFixed(2)} A ${outerR.toFixed(2)} ${outerR.toFixed(2)} 0 ${large} 1 ${x1.toFixed(2)} ${y1.toFixed(2)} Z`
  }
  const xi1 = cx + innerR * Math.cos(a1)
  const yi1 = cy + innerR * Math.sin(a1)
  const xi0 = cx + innerR * Math.cos(a0)
  const yi0 = cy + innerR * Math.sin(a0)
  return (
    `M ${x0.toFixed(2)} ${y0.toFixed(2)} A ${outerR.toFixed(2)} ${outerR.toFixed(2)} 0 ${large} 1 ${x1.toFixed(2)} ${y1.toFixed(2)} ` +
    `L ${xi1.toFixed(2)} ${yi1.toFixed(2)} A ${innerR.toFixed(2)} ${innerR.toFixed(2)} 0 ${large} 0 ${xi0.toFixed(2)} ${yi0.toFixed(2)} Z`
  )
}

/** one node (recursively) → SVG markup at (ox, oy) offset in parent space */
function nodeSvg(
  n: RenderNode,
  ox: number,
  oy: number,
  images: Map<string, HTMLImageElement>,
): string {
  const x = ox + n.box.x
  const y = oy + n.box.y
  const defs: string[] = []
  // rotation pivots the box center; geometry flips mirror about the center
  const parts: string[] = []
  if (n.box.rotationDeg) {
    parts.push(
      `rotate(${n.box.rotationDeg} ${(n.box.w / 2).toFixed(2)} ${(n.box.h / 2).toFixed(2)})`,
    )
  }
  const flipGeo =
    n.box.flipH || n.box.flipV
      ? `translate(${n.box.flipH ? n.box.w.toFixed(2) : 0} ${n.box.flipV ? n.box.h.toFixed(2) : 0}) scale(${n.box.flipH ? -1 : 1} ${n.box.flipV ? -1 : 1})`
      : ''

  let body = ''
  if (n.type === 'group') {
    body = (n as GroupRenderNode).children.map((c) => nodeSvg(c, 0, 0, images)).join('')
  } else if (n.type === 'picture') {
    const pic = n as PictureRenderNode
    const { box } = pic
    const src = pic.dataUrl
    if (src) {
      const c = pic.srcRect
      let img: string
      if (c && (c.l || c.t || c.r || c.b)) {
        // draw the visible source region into the frame (negative = outset)
        const vw = Math.max(1 - c.l - c.r, 0.01)
        const vh = Math.max(1 - c.t - c.b, 0.01)
        const iw = box.w / vw
        const ih = box.h / vh
        img = `<image href="${escAttr(src)}" x="${(-c.l * iw).toFixed(2)}" y="${(-c.t * ih).toFixed(2)}" width="${iw.toFixed(2)}" height="${ih.toFixed(2)}" preserveAspectRatio="none"`
      } else {
        img = `<image href="${escAttr(src)}" x="0" y="0" width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" preserveAspectRatio="none"`
      }
      if (pic.opacity !== undefined && pic.opacity < 1) img += ` opacity="${pic.opacity}"`
      img += '/>'
      // picture-style clips (rounded / geometric frames)
      let clipAttr = ''
      const clip = pic.clip
      if (clip?.pathData) {
        const id = `clip${defSeq++}`
        defs.push(`<clipPath id="${id}"><path d="${clip.pathData}"/></clipPath>`)
        clipAttr = ` clip-path="url(#${id})"`
      } else if (clip?.cornerRadiusPx) {
        const id = `clip${defSeq++}`
        defs.push(
          `<clipPath id="${id}"><rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" rx="${clip.cornerRadiusPx.toFixed(2)}"/></clipPath>`,
        )
        clipAttr = ` clip-path="url(#${id})"`
      }
      const bg = pic.bgColor
        ? `<rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" fill="${cssColor(pic.bgColor)}"/>`
        : ''
      const frame = pic.stroke
        ? `<rect width="${box.w.toFixed(2)}" height="${box.h.toFixed(2)}" fill="none"${strokeAttrs(pic.stroke)}/>`
        : ''
      body = `${bg}<g${clipAttr}>${img}</g>${frame}`
    }
  } else if (n.type === 'chart') {
    body = chartGeometry(n as ChartRenderNode, defs)
  } else if (n.type === 'table') {
    const table = n as TableRenderNode
    const { box } = table
    const tblW = table.gridX[table.gridX.length - 1] ?? box.w
    const tblH = table.gridY[table.gridY.length - 1] ?? box.h
    body =
      (table.bgFill && table.bgFill.kind !== 'none'
        ? `<rect width="${tblW.toFixed(2)}" height="${tblH.toFixed(2)}" fill="${paintFill(table.bgFill, tblW, tblH, defs).fill}"/>`
        : '') +
      table.cells
        .map((cell) => {
          const f = paintFill(cell.fill, cell.w, cell.h, defs)
          const b = cell.borders
          return (
            `<rect x="${cell.x.toFixed(2)}" y="${cell.y.toFixed(2)}" width="${cell.w.toFixed(2)}" height="${cell.h.toFixed(2)}" fill="${f.fill}"/>` +
            (b?.t
              ? `<line x1="${cell.x.toFixed(2)}" y1="${cell.y.toFixed(2)}" x2="${(cell.x + cell.w).toFixed(2)}" y2="${cell.y.toFixed(2)}"${strokeAttrs(b.t)}/>`
              : '') +
            (b?.b
              ? `<line x1="${cell.x.toFixed(2)}" y1="${(cell.y + cell.h).toFixed(2)}" x2="${(cell.x + cell.w).toFixed(2)}" y2="${(cell.y + cell.h).toFixed(2)}"${strokeAttrs(b.b)}/>`
              : '') +
            (b?.l
              ? `<line x1="${cell.x.toFixed(2)}" y1="${cell.y.toFixed(2)}" x2="${cell.x.toFixed(2)}" y2="${(cell.y + cell.h).toFixed(2)}"${strokeAttrs(b.l)}/>`
              : '') +
            (b?.r
              ? `<line x1="${(cell.x + cell.w).toFixed(2)}" y1="${cell.y.toFixed(2)}" x2="${(cell.x + cell.w).toFixed(2)}" y2="${(cell.y + cell.h).toFixed(2)}"${strokeAttrs(b.r)}/>`
              : '') +
            layoutText(cell.text, cell.x, cell.y)
          )
        })
        .join('')
  } else if (n.type === 'text' || n.type === 'shape') {
    const shape = n as ShapeRenderNode
    const geo = shapeGeometry(shape, defs)
    // PowerPoint flips geometry only; text mirrors position but stays readable
    const text = layoutText(shape.text, 0, 0, n.box.flipH ? n.box.w : 0, n.box.flipV ? n.box.h : 0)
    body = flipGeo
      ? `<g${flipGeo ? ` transform="${flipGeo}"` : ''}>${geo}</g>${text}`
      : `${geo}${text}`
  }

  if (parts.length) {
    return `<g transform="translate(${x.toFixed(2)} ${y.toFixed(2)}) ${parts.join(' ')}">${defsMarkup(defs)}${body}</g>`
  }
  return `<g transform="translate(${x.toFixed(2)} ${y.toFixed(2)})">${defsMarkup(defs)}${body}</g>`
}

const defsMarkup = (defs: string[]): string => (defs.length ? `<defs>${defs.join('')}</defs>` : '')

/**
 * Render one slide to a self-contained inline SVG (real text elements).
 * Throws on unexpected structures — callers fall back to the raster page.
 */
export function renderSlideSvg(slide: RenderSlide, images: Map<string, HTMLImageElement>): string {
  defSeq = 0
  const defs: string[] = []
  const bg = paintFill(slide.background, slide.widthPx, slide.heightPx, defs)
  const nodes = slide.nodes.map((n) => nodeSvg(n, 0, 0, images)).join('')
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${slide.widthPx} ${slide.heightPx}" width="100%" height="100%" preserveAspectRatio="none">` +
    defsMarkup(defs) +
    `<rect width="${slide.widthPx}" height="${slide.heightPx}" fill="${bg.fill}"/>` +
    nodes +
    `</svg>`
  )
}

/**
 * SmartArt insertion: generates the real OOXML Diagram parts for the four
 * insert presets — the dgm:dataModel (data1.xml), the MS pre-resolved layout
 * cache (dsp:drawing, drawing1.xml) and the paragraph fragment referencing the
 * quartet (dgm:relIds). layout/colors/quickStyle parts are vendored Office
 * built-ins (smartart-vendor.ts); Word embeds a copy of the chosen built-in
 * layout per saved diagram, so this matches Word's own package anatomy.
 *
 * Node modelIds are assigned once and shared between the data model and the
 * drawing cache: consumers correlate cached dsp:sp shapes to dgm:pt points by
 * modelId. The drawing cache doubles as the editor preview —
 * buildDiagramDisplay runs the same geometry and returns the DiagramDisplay
 * the parse path produces for saved files, so a just-inserted diagram renders
 * through the same display code before the first save.
 */
import { SMARTART_PRESET_PARTS } from './smartart-vendor'
import type { DiagramDisplay, DiagramShape, NewDiagram, NewDiagramItem } from './types'
import { escapeXmlText } from './xml-utils'

const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
const DIAGRAM_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
const DATA_MODEL_EXT_URI = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
const MIN_DIAGRAM_VER = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'

/** theme accent1 + the follow-on accents Word's colorful org-chart branches use */
const ACCENT1_HEX = '4472C4'
const BRANCH_HEXES = ['ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478']
const CONNECTOR_HEX = 'A5A5A5'
const BODY_TEXT_HEX = '404040'
/** 1pt stroke for the org-chart connectors */
const CONNECTOR_W_EMU = 12700

/** hard cap so a pathological caller cannot blow up the package (UI clamps at 8) */
const MAX_NODES = 64

/** Content tree node: level comes from the indented text pane (flat presets use level 0) */
interface TreeNode {
  text: string
  level: number
  children: TreeNode[]
}

/** one resolved shape in the diagram's own EMU coordinate space */
interface ShapeEmu {
  /** data point modelId this shape presents (decoration shapes get a synthetic id) */
  modelId: string
  x: number
  y: number
  cx: number
  cy: number
  prst: string
  fillHex?: string
  lineHex?: string
  text?: string
  /** run size in hundredths of a point (a:rPr sz) */
  sz: number
  bold?: boolean
  textColorHex?: string
  align?: 'ctr' | 'l'
}

function randomHexDigits(length: number): string {
  const cryptoApi = globalThis.crypto
  let out = ''
  if (cryptoApi?.getRandomValues) {
    const bytes = new Uint8Array(length)
    cryptoApi.getRandomValues(bytes)
    for (let i = 0; i < length; i++) out += (bytes[i]! % 16).toString(16)
  } else {
    for (let i = 0; i < length; i++) out += Math.floor(Math.random() * 16).toString(16)
  }
  return out
}

/** Fresh uppercase GUID in Office's braced modelId form ({XXXXXXXX-…}). */
export function freshDiagramGuid(): string {
  const h = (n: number) => randomHexDigits(n)
  return `{${h(8)}-${h(4)}-${h(4)}-${h(4)}-${h(12)}}`.toUpperCase()
}

/** Clamp node count, collapse invalid indent jumps, drop empties. */
export function normalizeDiagramItems(
  items: NewDiagramItem[],
): Array<{ text: string; level: number }> {
  const out: Array<{ text: string; level: number }> = []
  let prevLevel = 0
  for (const item of items) {
    if (out.length >= MAX_NODES) break
    const text = item.text.replace(/\s+/g, ' ').trim()
    if (!text) continue
    const raw = Math.max(0, Math.floor(item.level ?? 0))
    const level = out.length === 0 ? 0 : Math.min(raw, prevLevel + 1)
    prevLevel = level
    out.push({ text, level })
  }
  if (out.length === 0) out.push({ text: '[Text]', level: 0 })
  return out
}

/** Build the parent/child tree from indented items (flat presets produce a star from the root). */
function buildTree(items: Array<{ text: string; level: number }>, flat: boolean): TreeNode {
  const root: TreeNode = { text: '', level: -1, children: [] }
  if (flat) {
    for (const item of items) root.children.push({ text: item.text, level: 0, children: [] })
    return root
  }
  const stack: TreeNode[] = [root]
  for (const item of items) {
    const node: TreeNode = { text: item.text, level: item.level, children: [] }
    while (stack.length > item.level + 1) stack.pop()
    stack[stack.length - 1]!.children.push(node)
    stack.push(node)
  }
  return root
}

// ---------------------------------------------------------------------------
// geometry: one routine, two consumers (drawing1.xml EMU bodies + editor px preview)
// ---------------------------------------------------------------------------

/** Basic Block List: snake rows of rounded rectangles (accent1, white bold text) */
function blockListShapes(
  nodes: TreeNode[],
  ids: string[],
  cx: number,
): { shapes: ShapeEmu[]; cy: number } {
  const n = nodes.length
  const gap = 152400
  const cols = n <= 4 ? n : Math.ceil(n / 2)
  const rows = Math.ceil(n / cols)
  const w = Math.round((cx - gap * (cols - 1)) / cols)
  const h = Math.min(Math.max(Math.round(w * 0.62), 685800), 1066800)
  const shapes: ShapeEmu[] = nodes.map((node, i) => ({
    modelId: ids[i] ?? freshDiagramGuid(),
    x: (i % cols) * (w + gap),
    y: Math.floor(i / cols) * (h + gap),
    cx: w,
    cy: h,
    prst: 'roundRect',
    fillHex: ACCENT1_HEX,
    text: node.text,
    sz: 1400,
    bold: true,
    textColorHex: 'FFFFFF',
    align: 'ctr' as const,
  }))
  return { shapes, cy: rows * h + (rows - 1) * gap }
}

/** Vertical Bullet List: accent dots + left-aligned text rows, no boxes */
function vBulletShapes(
  nodes: TreeNode[],
  ids: string[],
  cx: number,
): { shapes: ShapeEmu[]; cy: number } {
  const rowH = 1143000
  const gap = 342900
  const dot = 228600
  const dotGap = 114300
  const shapes: ShapeEmu[] = []
  nodes.forEach((node, i) => {
    const y = i * (rowH + gap)
    shapes.push({
      modelId: freshDiagramGuid(),
      x: 0,
      y: y + Math.round((rowH - dot) / 2),
      cx: dot,
      cy: dot,
      prst: 'ellipse',
      fillHex: ACCENT1_HEX,
      sz: 100,
    })
    shapes.push({
      modelId: ids[i] ?? freshDiagramGuid(),
      x: dot + dotGap,
      y,
      cx: Math.max(0, cx - dot - dotGap),
      cy: rowH,
      prst: 'rect',
      text: node.text,
      sz: 1400,
      textColorHex: BODY_TEXT_HEX,
      align: 'l',
    })
  })
  return { shapes, cy: nodes.length * rowH + (nodes.length - 1) * gap }
}

/** Basic Process: interlocked chevrons flowing left to right */
function processShapes(nodes: TreeNode[], ids: string[], cx: number, cy: number): ShapeEmu[] {
  const n = nodes.length
  const overlap = 0.25
  const w = Math.round(cx / (n - (n - 1) * overlap))
  const h = Math.round(cy * 0.55)
  const y = Math.round((cy - h) / 2)
  return nodes.map((node, i) => ({
    modelId: ids[i] ?? freshDiagramGuid(),
    x: Math.round(i * w * (1 - overlap)),
    y,
    cx: w,
    cy: h,
    prst: i === 0 ? 'homePlate' : 'chevron',
    fillHex: ACCENT1_HEX,
    text: node.text,
    sz: 1400,
    bold: true,
    textColorHex: 'FFFFFF',
    align: 'ctr' as const,
  }))
}

/** Organization Chart: tidy tree, branch colors, elbow connectors */
function hierShapes(
  root: TreeNode,
  ids: Map<TreeNode, string>,
  cx: number,
): { shapes: ShapeEmu[]; cxOut: number; cy: number } {
  const w = 1371600
  const h = 914400
  const hGap = 274320
  const vGap = 914400
  const positions = new Map<TreeNode, { x: number; y: number }>()

  // pass 1: leaves left to right in reading order at a fixed slot pitch
  let leafCount = 0
  const assignLeaves = (node: TreeNode, level: number): void => {
    if (node.children.length === 0) {
      positions.set(node, { x: leafCount * (w + hGap), y: level * (h + vGap) })
      leafCount++
    } else {
      for (const child of node.children) assignLeaves(child, level + 1)
    }
  }
  for (const child of root.children) assignLeaves(child, 0)
  // pass 2: parents centered over their children (bottom-up)
  const centerParents = (node: TreeNode, level: number): void => {
    for (const child of node.children) centerParents(child, level + 1)
    if (node.children.length === 0) return
    const xs = node.children.map((c) => positions.get(c)!.x)
    const first = Math.min(...xs)
    const last = Math.max(...xs)
    positions.set(node, { x: Math.round((first + last + w) / 2) - w, y: level * (h + vGap) })
  }
  for (const child of root.children) centerParents(child, 0)

  const centerOf = (node: TreeNode): number => positions.get(node)!.x + Math.round(w / 2)
  const shapes: ShapeEmu[] = []
  const emitConnectors = (parent: TreeNode, children: TreeNode[]): void => {
    const parentPos = positions.get(parent)!
    const busY = parentPos.y + h + Math.round(vGap / 2)
    const line = (x: number, y: number, length: number, vertical: boolean): void => {
      shapes.push({
        modelId: freshDiagramGuid(),
        x,
        y,
        cx: vertical ? 0 : length,
        cy: vertical ? length : 0,
        prst: 'line',
        lineHex: CONNECTOR_HEX,
        sz: 100,
      })
    }
    line(centerOf(parent), parentPos.y + h, Math.round(vGap / 2), true)
    const childCenters = children.map(centerOf)
    const first = Math.min(centerOf(parent), ...childCenters)
    const last = Math.max(centerOf(parent), ...childCenters)
    if (last > first) line(first, busY, last - first, false)
    for (const child of children) {
      line(centerOf(child), busY, positions.get(child)!.y - busY, true)
    }
  }
  const emitNode = (node: TreeNode, level: number, hex: string): void => {
    const p = positions.get(node)!
    shapes.push({
      modelId: ids.get(node) ?? freshDiagramGuid(),
      x: p.x,
      y: p.y,
      cx: w,
      cy: h,
      prst: 'roundRect',
      fillHex: hex,
      text: node.text,
      sz: 1400,
      bold: true,
      textColorHex: 'FFFFFF',
      align: 'ctr',
    })
    emitConnectors(node, node.children)
    node.children.forEach((child, i) => {
      const childHex = level === 0 ? (BRANCH_HEXES[i % BRANCH_HEXES.length] ?? ACCENT1_HEX) : hex
      emitNode(child, level + 1, childHex)
    })
  }
  root.children.forEach((child, i) => {
    const hex = i === 0 ? ACCENT1_HEX : (BRANCH_HEXES[(i - 1) % BRANCH_HEXES.length] ?? ACCENT1_HEX)
    emitNode(child, 0, hex)
  })

  let maxLevel = 0
  for (const p of positions.values()) maxLevel = Math.max(maxLevel, Math.round(p.y / (h + vGap)))
  const cy = (maxLevel + 1) * h + maxLevel * vGap
  const totalW = leafCount * w + (leafCount - 1) * hGap
  return { shapes, cxOut: Math.max(cx, totalW), cy }
}

/** default insert extents (EMU) per preset — Word-like gallery proportions */
export function diagramDefaultExtentEmu(spec: Pick<NewDiagram, 'kind' | 'items'>): {
  cx: number
  cy: number
} {
  const items = normalizeDiagramItems(spec.items)
  const root = buildTree(items, spec.kind !== 'hier')
  switch (spec.kind) {
    case 'blockList':
      return { cx: 5486400, cy: blockListShapes(root.children, [], 5486400).cy }
    case 'vBulletList':
      return { cx: 3657600, cy: vBulletShapes(root.children, [], 3657600).cy }
    case 'process':
      return { cx: 5486400, cy: 1828800 }
    case 'hier': {
      const { cxOut, cy } = hierShapes(root, new Map(), 4572000)
      return { cx: cxOut, cy }
    }
  }
}

/**
 * Assign the data-point modelIds once so data1.xml points and drawing1.xml
 * dsp:sp shapes correlate. Decoration shapes (bullet dots, connector rules)
 * take synthetic ids of their own.
 */
function assignNodeIds(root: TreeNode): Map<TreeNode, string> {
  const ids = new Map<TreeNode, string>()
  const visit = (node: TreeNode): void => {
    ids.set(node, freshDiagramGuid())
    for (const child of node.children) visit(child)
  }
  for (const child of root.children) visit(child)
  return ids
}

/** Resolve the shapes for a tree (EMU space) against the shared node model ids. */
function diagramShapes(
  kind: NewDiagram['kind'],
  root: TreeNode,
  extent: { cx: number; cy: number },
  ids: Map<TreeNode, string>,
): ShapeEmu[] {
  const flatIds: string[] = []
  const collect = (node: TreeNode): void => {
    flatIds.push(ids.get(node) ?? freshDiagramGuid())
    for (const child of node.children) collect(child)
  }
  for (const child of root.children) collect(child)

  switch (kind) {
    case 'blockList':
      return blockListShapes(root.children, flatIds, extent.cx).shapes
    case 'vBulletList':
      return vBulletShapes(root.children, flatIds, extent.cx).shapes
    case 'process':
      return processShapes(root.children, flatIds, extent.cx, extent.cy)
    case 'hier':
      return hierShapes(root, ids, extent.cx).shapes
  }
}

// ---------------------------------------------------------------------------
// part XML builders
// ---------------------------------------------------------------------------

function runXml(shape: {
  text?: string
  sz: number
  bold?: boolean
  textColorHex?: string
}): string {
  return (
    `<a:rPr lang="en-US" sz="${shape.sz}"${shape.bold ? ' b="1"' : ''}` +
    (shape.textColorHex
      ? `><a:solidFill><a:srgbClr val="${shape.textColorHex}"/></a:solidFill></a:rPr>`
      : '/>') +
    `<a:t>${escapeXmlText(shape.text ?? '')}</a:t>`
  )
}

function txBodyXml(shape: ShapeEmu): string {
  const align = shape.align === 'l' ? '' : '<a:pPr algn="ctr"/>'
  return (
    `<dsp:txBody>` +
    '<a:bodyPr wrap="square" lIns="45720" tIns="27432" rIns="45720" bIns="27432" anchor="ctr"/>' +
    `<a:lstStyle/><a:p>${align}<a:r>${runXml(shape)}</a:r></a:p>` +
    `</dsp:txBody>`
  )
}

function drawingShapeXml(shape: ShapeEmu, id: number): string {
  const fill = shape.fillHex
    ? `<a:solidFill><a:srgbClr val="${shape.fillHex}"/></a:solidFill>`
    : '<a:noFill/>'
  const ln = shape.lineHex
    ? `<a:ln w="${CONNECTOR_W_EMU}"><a:solidFill><a:srgbClr val="${shape.lineHex}"/></a:solidFill></a:ln>`
    : '<a:ln><a:noFill/></a:ln>'
  return (
    `<dsp:sp modelId="${shape.modelId}">` +
    `<dsp:nvSpPr><dsp:cNvPr id="${id}" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>` +
    `<dsp:spPr>` +
    `<a:xfrm><a:off x="${shape.x}" y="${shape.y}"/><a:ext cx="${shape.cx}" cy="${shape.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${shape.prst}"><a:avLst/></a:prstGeom>` +
    `${fill}${ln}` +
    `</dsp:spPr>${shape.text ? txBodyXml(shape) : ''}</dsp:sp>`
  )
}

/**
 * The generated dgm:dataModel part. Presentation/transition points are Word's
 * layout-run cache and stay away (Word regenerates them by re-running the
 * layout engine); connections use the ISO §21.4 minimal parOf form.
 */
function dataModelXml(
  kind: NewDiagram['kind'],
  root: TreeNode,
  ids: Map<TreeNode, string>,
  drawingRelId: string,
): string {
  const preset = SMARTART_PRESET_PARTS[kind]
  const docId = freshDiagramGuid()
  const points: string[] = [
    `<dgm:pt modelId="${docId}" type="doc">` +
      `<dgm:prSet loTypeId="${preset.loTypeId}" loCatId="${preset.loCatId}"` +
      ` qsTypeId="${preset.qsTypeId}" qsCatId="${preset.qsCatId}"` +
      ` csTypeId="${preset.csTypeId}" csCatId="${preset.csCatId}" phldr="1"/>` +
      '<dgm:spPr/></dgm:pt>',
  ]
  const cxns: string[] = []
  const registerPoint = (node: TreeNode): void => {
    points.push(
      `<dgm:pt modelId="${ids.get(node)!}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/>` +
        `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>` +
        `<a:t>${escapeXmlText(node.text)}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
    )
    for (const child of node.children) registerPoint(child)
  }
  const registerCxns = (node: TreeNode): void => {
    node.children.forEach((child, i) => {
      cxns.push(
        `<dgm:cxn modelId="${freshDiagramGuid()}" srcId="${ids.get(node)!}"` +
          ` destId="${ids.get(child)!}" srcOrd="${i}" destOrd="0"/>`,
      )
      registerCxns(child)
    })
  }
  root.children.forEach((child) => registerPoint(child))
  root.children.forEach((child, i) => {
    cxns.push(
      `<dgm:cxn modelId="${freshDiagramGuid()}" srcId="${docId}"` +
        ` destId="${ids.get(child)!}" srcOrd="${i}" destOrd="0"/>`,
    )
    registerCxns(child)
  })

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<dgm:dataModel xmlns:dgm="${DGM_NS}" xmlns:a="${A_NS}">` +
    `<dgm:ptLst>${points.join('')}</dgm:ptLst>` +
    `<dgm:cxnLst>${cxns.join('')}</dgm:cxnLst>` +
    '<dgm:bg/><dgm:whole/>' +
    `<dgm:extLst><a:ext uri="${DATA_MODEL_EXT_URI}">` +
    `<dsp:dataModelExt xmlns:dsp="${DSP_NS}" relId="${drawingRelId}" minVer="${MIN_DIAGRAM_VER}"/>` +
    '</a:ext></dgm:extLst>' +
    '</dgm:dataModel>'
  )
}

/** The generated MS extension part (dsp:drawing): our resolved geometry. */
function drawingXml(shapes: ShapeEmu[], extent: { cx: number; cy: number }): string {
  const body = shapes.map((shape, i) => drawingShapeXml(shape, i + 1)).join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}">` +
    '<dsp:spTree>' +
    '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>' +
    `<dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${extent.cx}" cy="${extent.cy}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${extent.cx}" cy="${extent.cy}"/></a:xfrm></dsp:grpSpPr>` +
    `${body}</dsp:spTree></dsp:drawing>`
  )
}

/** The generated diagram pair (data1.xml + drawing1.xml bodies). */
export function buildDiagramPartsXml(
  spec: NewDiagram,
  opts: { drawingRelId: string; extent: { cx: number; cy: number } },
): { dataXml: string; drawingXml: string } {
  const items = normalizeDiagramItems(spec.items)
  const root = buildTree(items, spec.kind !== 'hier')
  const ids = assignNodeIds(root)
  return {
    dataXml: dataModelXml(spec.kind, root, ids, opts.drawingRelId),
    drawingXml: drawingXml(diagramShapes(spec.kind, root, opts.extent, ids), opts.extent),
  }
}

/**
 * Editor preview for a not-yet-saved insert: the same geometry as the drawing
 * part, in the px DiagramDisplay form the parser produces for saved files.
 */
export function buildDiagramDisplay(
  spec: NewDiagram,
  extent?: { cx: number; cy: number },
): DiagramDisplay {
  const ext = extent ?? diagramDefaultExtentEmu(spec)
  const root = buildTree(normalizeDiagramItems(spec.items), spec.kind !== 'hier')
  const shapes = diagramShapes(spec.kind, root, ext, assignNodeIds(root))
  const px = (v: number) => Math.round(v / 9525)
  const out: DiagramShape[] = []
  for (const shape of shapes) {
    const display: DiagramShape = {
      xPx: px(shape.x),
      yPx: px(shape.y),
      wPx: px(shape.cx),
      hPx: px(shape.cy),
    }
    if (shape.prst !== 'rect') display.prst = shape.prst
    if (shape.fillHex) display.fillHex = shape.fillHex
    if (shape.lineHex) {
      display.lnHex = shape.lineHex
      display.lnWPx = Math.max(1, Math.round(CONNECTOR_W_EMU / 9525))
    }
    if (shape.text) {
      display.texts = [shape.text]
      display.fontSizePt = Math.round(shape.sz / 100)
      if (shape.textColorHex) display.textColorHex = shape.textColorHex
    }
    out.push(display)
  }
  return { widthPx: px(ext.cx), heightPx: px(ext.cy), shapes: out }
}

/** The w:p fragment referencing the diagram quartet (dgm:relIds). */
export function diagramParagraphXml(opts: {
  docPrId: number
  extent: { cx: number; cy: number }
  dmRId: string
  loRId: string
  qsRId: string
  csRId: string
}): string {
  const { docPrId, extent } = opts
  return (
    '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${extent.cx}" cy="${extent.cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${docPrId}" name="Diagram ${docPrId}"/>` +
    '<wp:cNvGraphicFramePr/>' +
    `<a:graphic xmlns:a="${A_NS}">` +
    `<a:graphicData uri="${DIAGRAM_GRAPHIC_URI}">` +
    `<dgm:relIds xmlns:dgm="${DGM_NS}" xmlns:r="${R_NS}"` +
    ` r:dm="${opts.dmRId}" r:lo="${opts.loRId}" r:qs="${opts.qsRId}" r:cs="${opts.csRId}"/>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
  )
}

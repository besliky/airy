import JSZip from 'jszip'
import { XMLValidator } from 'fast-xml-parser'
import { describe, expect, it } from 'vitest'
import {
  buildDiagramDisplay,
  decodeEntities,
  parseDocx,
  saveDocx,
  SMARTART_PRESET_PARTS,
  type NewDiagram,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const FLAT_ITEMS = ['调研', '设计 <验收>', '开发', '发布']

const PRESETS: NewDiagram[] = [
  { kind: 'blockList', items: FLAT_ITEMS.map((text) => ({ text })) },
  { kind: 'vBulletList', items: FLAT_ITEMS.map((text) => ({ text })) },
  { kind: 'process', items: FLAT_ITEMS.map((text) => ({ text })) },
  {
    kind: 'hier',
    items: [
      { text: '总经理', level: 0 },
      { text: '研发部', level: 1 },
      { text: '平台组', level: 2 },
      { text: '销售部', level: 1 },
      { text: '运营部', level: 1 },
    ],
  },
]

const DIAGRAM_PART_CONTENT_TYPES: Record<string, string> = {
  data: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
  layout: 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
  quickStyle: 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
  colors: 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
  drawing: 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
}

async function insertDiagram(
  spec: NewDiagram,
  options: { existingDataPart?: boolean } = {},
): Promise<JSZip> {
  const bodyXml = options.existingDataPart
    ? '<w:p><w:r><w:t>src</w:t></w:r></w:p>' +
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
      '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
      'r:dm="rId40" r:lo="rId41" r:qs="rId42" r:cs="rId43"/>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    : '<w:p><w:r><w:t>前文</w:t></w:r></w:p>'
  const source = await buildDocx({
    bodyXml,
    extraRels: options.existingDataPart
      ? '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>'
      : undefined,
    extraParts: options.existingDataPart
      ? [
          {
            path: 'word/diagrams/data1.xml',
            xml:
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>' +
              '<dgm:pt modelId="{1}"><dgm:t><a:bodyPr/><a:p><a:r><a:t>既有</a:t></a:r></a:p></dgm:t></dgm:pt>' +
              '</dgm:ptLst></dgm:dataModel>',
            contentType: DIAGRAM_PART_CONTENT_TYPES.data!,
          },
        ]
      : undefined,
  })
  const parsed = await parseDocx(source)
  const blocks: Array<
    { kind: 'original'; docxIndex: number } | { kind: 'diagram'; diagram: NewDiagram }
  > = parsed.blocks
    .filter((b) => !b.hidden)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
  blocks.push({ kind: 'diagram', diagram: spec })
  const saved = await saveDocx(parsed, blocks)
  return JSZip.loadAsync(saved)
}

function pointTexts(dataXml: string): string[] {
  const out: string[] = []
  for (const m of dataXml.matchAll(/<dgm:pt modelId="[^"]+"[^>]*>[\s\S]*?<\/dgm:pt>/g)) {
    const text = /<a:t>([^<]*)<\/a:t>/.exec(m[0])?.[1]
    if (text !== undefined) out.push(decodeEntities(text))
  }
  return out
}

function pointModelIds(dataXml: string): string[] {
  return [...dataXml.matchAll(/<dgm:pt modelId="([^"]+)"/g)].map((m) => m[1]!)
}

describe('saveDocx kind:diagram — package anatomy per preset', () => {
  for (const spec of PRESETS) {
    it(`writes the five-part quartet for ${spec.kind}`, async () => {
      const zip = await insertDiagram(spec)
      const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
      const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')

      const dataXml = await zip.file('word/diagrams/data1.xml')!.async('string')
      const layoutXml = await zip.file('word/diagrams/layout1.xml')!.async('string')
      const colorsXml = await zip.file('word/diagrams/colors1.xml')!.async('string')
      const quickStyleXml = await zip.file('word/diagrams/quickStyle1.xml')!.async('string')
      const drawingXml = await zip.file('word/diagrams/drawing1.xml')!.async('string')

      // content types: the asymmetric strings are the classic repair traps
      expect(contentTypes).toContain(
        `<Override PartName="/word/diagrams/data1.xml" ContentType="${DIAGRAM_PART_CONTENT_TYPES.data}"/>`,
      )
      expect(contentTypes).toContain(
        `<Override PartName="/word/diagrams/layout1.xml" ContentType="${DIAGRAM_PART_CONTENT_TYPES.layout}"/>`,
      )
      expect(contentTypes).toContain(
        `<Override PartName="/word/diagrams/quickStyle1.xml" ContentType="${DIAGRAM_PART_CONTENT_TYPES.quickStyle}"/>`,
      )
      expect(contentTypes).toContain(
        `<Override PartName="/word/diagrams/colors1.xml" ContentType="${DIAGRAM_PART_CONTENT_TYPES.colors}"/>`,
      )
      expect(contentTypes).toContain(
        `<Override PartName="/word/diagrams/drawing1.xml" ContentType="${DIAGRAM_PART_CONTENT_TYPES.drawing}"/>`,
      )

      // relationships: 4 standard diagram rels + the MS diagramDrawing rel
      expect(rels).toContain('relationships/diagramData" Target="diagrams/data1.xml"')
      expect(rels).toContain('relationships/diagramLayout" Target="diagrams/layout1.xml"')
      expect(rels).toContain('relationships/diagramQuickStyle" Target="diagrams/quickStyle1.xml"')
      expect(rels).toContain('relationships/diagramColors" Target="diagrams/colors1.xml"')
      expect(rels).toContain(
        'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"',
      )

      // vendored parts land byte-for-byte; doc point ids match their uniqueIds
      const preset = SMARTART_PRESET_PARTS[spec.kind]
      expect(layoutXml).toBe(preset.layoutXml)
      expect(colorsXml).toBe(preset.colorsXml)
      expect(quickStyleXml).toBe(preset.quickStyleXml)
      expect(dataXml).toContain(`loTypeId="${preset.loTypeId}"`)
      expect(dataXml).toContain(`qsTypeId="${preset.qsTypeId}"`)
      expect(dataXml).toContain(`csTypeId="${preset.csTypeId}"`)

      // data model: minimal parOf cxns, no presentation points
      expect(dataXml).not.toContain('type="pres"')
      expect(dataXml).not.toContain('parTrans')
      const texts = pointTexts(dataXml)
      expect(texts).toEqual(spec.items.map((i) => i.text))
      // every cxn is the attribute-minimal form
      for (const cxn of dataXml.match(/<dgm:cxn [^>]*\/>/g) ?? []) {
        expect(cxn).toMatch(
          /^<dgm:cxn modelId="[0-9A-F-{}]+" srcId="[^"]+" destId="[^"]+" srcOrd="\d+" destOrd="\d+"\/>$/,
        )
      }
      // the drawing-cache pointer: dataModelExt relId resolves in document rels
      const extRelId = /<dsp:dataModelExt[^>]*relId="([^"]+)"/.exec(dataXml)?.[1]
      expect(extRelId).toBeTruthy()
      expect(rels).toContain(`Id="${extRelId}"`)

      // drawing cache: dsp:sp with data-point modelIds + well-formed XML everywhere
      const shapeIds = [...drawingXml.matchAll(/<dsp:sp modelId="([^"]+)"/g)].map((m) => m[1]!)
      const ids = new Set(pointModelIds(dataXml))
      const textShapeIds = drawingXml
        .split('<dsp:sp modelId="')
        .slice(1)
        .filter((chunk) => chunk.includes('<dsp:txBody>'))
        .map((chunk) => chunk.slice(0, chunk.indexOf('"')))
      expect(textShapeIds.length).toBe(spec.items.length)
      for (const id of textShapeIds) expect(ids.has(id)).toBe(true)
      expect(shapeIds.length).toBeGreaterThan(0)
      for (const xml of [dataXml, layoutXml, colorsXml, quickStyleXml, drawingXml]) {
        expect(XMLValidator.validate(xml)).toBe(true)
      }
    })
  }

  it('hierarchy cxns encode the tree (doc → root → children → grandchildren)', async () => {
    const zip = await insertDiagram(PRESETS[3]!)
    const dataXml = await zip.file('word/diagrams/data1.xml')!.async('string')
    const idsByText = new Map(
      pointTexts(dataXml).map((t, i) => [t, pointModelIds(dataXml)[i + 1]!]),
    )
    const docId = /<dgm:pt modelId="([^"]+)" type="doc">/.exec(dataXml)?.[1]
    const cxns = [...dataXml.matchAll(/<dgm:cxn [^>]*\/>/g)].map((m) => m[0])
    const edgeOf = (from: string, to: string) =>
      cxns.find((c) => c.includes(`srcId="${from}"`) && c.includes(`destId="${to}"`))
    // doc -> general manager -> each department; R&D dept -> platform group (CJK node texts below are string data)
    expect(edgeOf(docId!, idsByText.get('总经理')!)).toBeTruthy()
    expect(edgeOf(idsByText.get('总经理')!, idsByText.get('研发部')!)).toBeTruthy()
    expect(edgeOf(idsByText.get('总经理')!, idsByText.get('销售部')!)).toBeTruthy()
    expect(edgeOf(idsByText.get('研发部')!, idsByText.get('平台组')!)).toBeTruthy()
    // sibling order encoded as srcOrd on the CEO edges
    const ceoCxns = cxns.filter((c) => c.includes(`srcId="${idsByText.get('总经理')!}"`))
    expect(ceoCxns.map((c) => /srcOrd="(\d+)"/.exec(c)?.[1])).toEqual(['0', '1', '2'])
  })

  it('allocates the next free shared index when diagrams/data1.xml already exists', async () => {
    const zip = await insertDiagram(PRESETS[0]!, { existingDataPart: true })
    // pre-existing data1.xml untouched; the new quartet takes index 2 in ALL five parts
    expect(await zip.file('word/diagrams/data1.xml')!.async('string')).toContain('既有')
    for (const name of ['data', 'layout', 'colors', 'quickStyle', 'drawing']) {
      expect(zip.file(`word/diagrams/${name}2.xml`), name).toBeTruthy()
      expect(zip.file(`word/diagrams/${name}3.xml`), name).toBeNull()
    }
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).toContain('/word/diagrams/drawing2.xml')
  })

  it('gives two diagrams in one save distinct paired indices', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'diagram', diagram: PRESETS[0]! },
      { kind: 'diagram', diagram: PRESETS[2]! },
    ])
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/diagrams/data1.xml')).toBeTruthy()
    expect(zip.file('word/diagrams/drawing1.xml')).toBeTruthy()
    expect(zip.file('word/diagrams/data2.xml')).toBeTruthy()
    expect(zip.file('word/diagrams/drawing2.xml')).toBeTruthy()
    const data1 = await zip.file('word/diagrams/data1.xml')!.async('string')
    const data2 = await zip.file('word/diagrams/data2.xml')!.async('string')
    expect(data1).toContain(SMARTART_PRESET_PARTS.blockList.loTypeId)
    expect(data2).toContain(SMARTART_PRESET_PARTS.process.loTypeId)
  })
})

describe('diagram round-trip', () => {
  it('re-parses as a SmartArt passthrough with a rendered diagramDisplay', async () => {
    for (const spec of PRESETS) {
      const zip = await insertDiagram(spec)
      const bytes = await zip.generateAsync({ type: 'uint8array' })
      const reparsed = await parseDocx(bytes)
      const block = reparsed.blocks.find((b) => b.diagramDisplay)
      expect(block, spec.kind).toBeDefined()
      expect(block!.type).toBe('passthrough')
      expect(block!.label).toBe('SmartArt')
      expect(block!.previewText).toBe(spec.items.map((i) => i.text).join('\n'))
      // flat presets carry one shape per node; the hierarchy adds connector rules
      expect(block!.diagramDisplay!.shapes.length).toBeGreaterThanOrEqual(spec.items.length)
      expect(block!.diagramDisplay!.widthPx).toBeGreaterThan(0)
    }
  })

  it('a no-edit second save keeps every diagram part byte-identical', async () => {
    const zip = await insertDiagram(PRESETS[3]!)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const parsed = await parseDocx(bytes)
    const saved2 = await saveDocx(
      parsed,
      parsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    const zip2 = await JSZip.loadAsync(saved2)
    for (const name of Object.keys(DIAGRAM_PART_CONTENT_TYPES)) {
      const before = await zip.file(`word/diagrams/${name}1.xml`)!.async('uint8array')
      const after = await zip2.file(`word/diagrams/${name}1.xml`)!.async('uint8array')
      expect(after, name).toEqual(before)
    }
    // untouched non-diagram part stays byte-identical too
    expect(await zip2.file('word/styles.xml')!.async('uint8array')).toEqual(
      await zip.file('word/styles.xml')!.async('uint8array'),
    )
  })

  it('deleting the diagram paragraph prunes all five parts, their rels and overrides', async () => {
    const zip = await insertDiagram(PRESETS[0]!)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const parsed = await parseDocx(bytes)
    const kept = parsed.blocks.filter((b) => !b.hidden && !b.diagramDisplay)
    const saved2 = await saveDocx(
      parsed,
      kept.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    const zip2 = await JSZip.loadAsync(saved2)
    for (const name of ['data', 'layout', 'colors', 'quickStyle', 'drawing']) {
      expect(zip2.file(`word/diagrams/${name}1.xml`), name).toBeNull()
    }
    const rels = await zip2.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).not.toContain('diagrams/')
    const contentTypes = await zip2.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).not.toContain('diagrams/')
  })
})

describe('buildDiagramDisplay (insert-time preview)', () => {
  it('mirrors the geometry the drawing part will carry', () => {
    const display = buildDiagramDisplay(PRESETS[3]!)
    const texts = display.shapes.filter((s) => s.texts?.length).map((s) => s.texts![0])
    expect(texts).toEqual(PRESETS[3]!.items.map((i) => i.text))
    // org chart carries elbow connector rules (zero width or height)
    const connectors = display.shapes.filter((s) => s.lnHex && (s.wPx <= 0 || s.hPx <= 0))
    expect(connectors.length).toBeGreaterThan(0)
    expect(display.widthPx).toBeGreaterThan(0)
    expect(display.heightPx).toBeGreaterThan(0)
  })

  it('escapes nothing into the display path and clamps wild inputs', () => {
    const display = buildDiagramDisplay({
      kind: 'blockList',
      items: [
        { text: 'a & b <c>', level: 3 },
        { text: '   ', level: 0 },
        { text: 'ok', level: 2 },
      ],
    })
    expect(display.shapes.filter((s) => s.texts?.length).map((s) => s.texts![0])).toEqual([
      'a & b <c>',
      'ok',
    ])
  })
})

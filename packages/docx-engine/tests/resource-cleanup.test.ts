import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { IMAGE_PARAGRAPH_XML, TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

function originalBlocksWithout(
  parsed: Awaited<ReturnType<typeof parseDocx>>,
  excluded: (block: Awaited<ReturnType<typeof parseDocx>>['blocks'][number]) => boolean,
): SaveBlock[] {
  return parsed.blocks
    .filter((block) => !block.hidden && !excluded(block))
    .map((block) => ({ kind: 'original', docxIndex: block.docxIndex! }))
}

async function documentRels(zip: JSZip): Promise<string> {
  return zip.file('word/_rels/document.xml.rels')!.async('string')
}

function mediaPaths(zip: JSZip): string[] {
  return Object.keys(zip.files).filter(
    (path) => /^word\/media\/[^/]+$/.test(path) && !zip.files[path].dir,
  )
}

describe('DOCX-owned resource cleanup', () => {
  it('removes the relationship and media part when an image block is deleted', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true }),
    )
    const zip = await JSZip.loadAsync(await saveDocx(parsed, []))
    expect(zip.file('word/media/image1.png')).toBeNull()
    expect(await documentRels(zip)).not.toContain('media/image1.png')
  })

  it('keeps only the newest media part across repeated image replacements', async () => {
    let bytes = await buildDocx({ bodyXml: IMAGE_PARAGRAPH_XML, withImage: true })
    for (let replacement = 0; replacement < 3; replacement++) {
      const parsed = await parseDocx(bytes)
      const image = parsed.blocks.find((block) => block.type === 'image')!
      bytes = await saveDocx(parsed, [
        {
          kind: 'xml',
          xml: image.originalXml!,
          docxIndex: image.docxIndex!,
          replaceImage: { base64: TINY_PNG_BASE64, mime: 'image/png' },
        },
      ])
      const zip = await JSZip.loadAsync(bytes)
      expect(mediaPaths(zip)).toHaveLength(1)
      expect(
        (await documentRels(zip)).match(new RegExp(`Type="${IMAGE_REL_TYPE}"`, 'g')),
      ).toHaveLength(1)
    }
    const finalZip = await JSZip.loadAsync(bytes)
    expect(mediaPaths(finalZip)).toEqual(['word/media/aidocs3.png'])
  })

  it('retains a shared image target while another surviving block still references it', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: IMAGE_PARAGRAPH_XML + IMAGE_PARAGRAPH_XML,
        withImage: true,
      }),
    )
    const zip = await JSZip.loadAsync(await saveDocx(parsed, [{ kind: 'original', docxIndex: 1 }]))
    expect(zip.file('word/media/image1.png')).not.toBeNull()
    expect(await documentRels(zip)).toContain('media/image1.png')
  })

  it('retains an image target referenced by a surviving header part', async () => {
    const sourceZip = await JSZip.loadAsync(
      await buildDocx({
        bodyXml: IMAGE_PARAGRAPH_XML,
        withImage: true,
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
        extraParts: [
          {
            path: 'word/header1.xml',
            xml:
              '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
              IMAGE_PARAGRAPH_XML +
              '</w:hdr>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
          },
        ],
        sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
      }),
    )
    sourceZip.file(
      'word/_rels/header1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId10" Type="${IMAGE_REL_TYPE}" Target="media/image1.png"/>` +
        '</Relationships>',
    )
    const parsed = await parseDocx(await sourceZip.generateAsync({ type: 'uint8array' }))
    const zip = await JSZip.loadAsync(
      await saveDocx(
        parsed,
        originalBlocksWithout(parsed, (block) => block.type === 'image'),
      ),
    )
    expect(zip.file('word/media/image1.png')).not.toBeNull()
    expect(await documentRels(zip)).not.toContain('media/image1.png')
    expect(await zip.file('word/_rels/header1.xml.rels')!.async('string')).toContain(
      'media/image1.png',
    )
  })

  it('removes an external hyperlink relationship after its link is removed', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:hyperlink r:id="rId20"><w:r><w:t>linked</w:t></w:r></w:hyperlink>' +
          '<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>',
        extraRels:
          '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
          'Target="https://example.com/" TargetMode="External"/>',
      }),
    )
    const zip = await JSZip.loadAsync(
      await saveDocx(parsed, [
        {
          kind: 'generated',
          block: { type: 'paragraph', runs: [{ text: 'linked' }, { text: ' tail' }] },
        },
      ]),
    )
    const rels = await documentRels(zip)
    expect(rels).not.toContain('relationships/hyperlink')
    expect(rels).not.toContain('https://example.com/')
  })

  it('removes a deleted chart and its relationship/workbook subgraph and overrides', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>keep</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(source)
    const withChart = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'chart',
        chart: {
          kind: 'bar',
          title: 'Sales',
          categories: ['Q1', 'Q2'],
          series: [{ name: 'East', values: [1, 2] }],
        },
      },
    ])
    const chartDoc = await parseDocx(withChart)
    const saved = await saveDocx(
      chartDoc,
      originalBlocksWithout(chartDoc, (block) => block.chartDisplay !== undefined),
    )
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/charts/chart1.xml')).toBeNull()
    expect(zip.file('word/charts/_rels/chart1.xml.rels')).toBeNull()
    expect(zip.file('word/charts/embeddings/workbook1.xlsx')).toBeNull()
    expect(await documentRels(zip)).not.toContain('relationships/chart')
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).not.toContain('/word/charts/chart1.xml')
    expect(contentTypes).not.toContain('/word/charts/embeddings/workbook1.xlsx')
  })

  it('cleans image resources owned by a deleted table block', async () => {
    const tableWithImage =
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:drawing><wp:inline>' +
      '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/>' +
      '</pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
      '</wp:inline></w:drawing></w:r></w:p></w:tc></w:tr></w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: tableWithImage, withImage: true }))
    expect(parsed.blocks[0].type).toBe('table')
    const zip = await JSZip.loadAsync(await saveDocx(parsed, []))
    expect(zip.file('word/media/image1.png')).toBeNull()
    expect(await documentRels(zip)).not.toContain('media/image1.png')
  })

  it('keeps the drawing-cache rel for a foreign-prefixed single-quoted dataModelExt (BUG-713)', async () => {
    // a non-Word producer may declare the Microsoft diagram namespace with a
    // prefix other than dsp: and single-quote its attributes — the retention
    // pass must still see the drawing-cache pointer, or the diagramDrawing
    // rel + drawing part get pruned while data1.xml still references them
    const bodyXml =
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
      '<dgm:relIds r:dm="rId40" r:lo="rId41" r:qs="rId42" r:cs="rId43"/>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    const dataXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
      'xmlns:zz="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
      '<dgm:ptLst><dgm:pt modelId="{1}" type="doc"/></dgm:ptLst>' +
      "<dgm:extLst><zz:dataModelExt relId='rId45' minVer='14.0'/></dgm:extLst>" +
      '</dgm:dataModel>'
    const drawingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
      '<dsp:spTree><dsp:nvGrpSpPr/><dsp:sp modelId="{1}"/></dsp:spTree></dsp:drawing>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml,
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>' +
          '<Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="diagrams/layout1.xml"/>' +
          '<Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>' +
          '<Relationship Id="rId43" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="diagrams/colors1.xml"/>' +
          '<Relationship Id="rId45" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>',
        extraParts: [
          {
            path: 'word/diagrams/data1.xml',
            xml: dataXml,
            contentType: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
          },
          {
            path: 'word/diagrams/layout1.xml',
            xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:layoutDef/>',
            contentType:
              'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
          },
          {
            path: 'word/diagrams/quickStyle1.xml',
            xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:styleDef/>',
            contentType: 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
          },
          {
            path: 'word/diagrams/colors1.xml',
            xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:colorsDef/>',
            contentType:
              'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
          },
          {
            path: 'word/diagrams/drawing1.xml',
            xml: drawingXml,
            contentType: 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
          },
        ],
      }),
    )
    // one edited paragraph, so the save runs the patch + cleanup pass (a
    // fully unchanged plan returns the original bytes and never cleans)
    const plan = originalBlocksWithout(parsed, () => false)
    plan.push({
      kind: 'generated',
      block: { type: 'paragraph', runs: [{ text: 'edited' }] },
    })
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan))
    // the diagram survives the save: data part, drawing part and every rel
    expect(zip.file('word/diagrams/data1.xml')).not.toBeNull()
    expect(zip.file('word/diagrams/drawing1.xml')).not.toBeNull()
    const rels = await documentRels(zip)
    expect(rels).toContain('Id="rId40"')
    expect(rels).toContain('Id="rId45"')
    expect(rels).toContain('relationships/diagramDrawing" Target="diagrams/drawing1.xml"')
  })

  it('cleans explicit OLE and diagram resources but preserves unknown custom parts', async () => {
    const bodyXml =
      '<w:p><w:r><w:object><v:shape><v:imagedata r:id="rId10"/></v:shape>' +
      '<o:OLEObject r:id="rId40" Type="Embed" ProgID="Package"/></w:object></w:r></w:p>' +
      '<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData>' +
      '<dgm:relIds r:dm="rId41"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
      '<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData>' +
      '<c:chart r:id="rId42"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml,
        withImage: true,
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>' +
          '<Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>' +
          '<Relationship Id="rId42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../customXml/custom-chart.xml"/>',
        extraParts: [
          {
            path: 'word/diagrams/data1.xml',
            xml: '<dgm:dataModel/>',
            contentType: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
          },
          {
            path: 'customXml/custom-chart.xml',
            xml: '<custom:extension/>',
            contentType: 'application/vnd.example.custom+xml',
          },
        ],
        binaryParts: [
          {
            path: 'word/embeddings/oleObject1.bin',
            base64: TINY_PNG_BASE64,
            extension: 'bin',
            contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
          },
        ],
      }),
    )
    const zip = await JSZip.loadAsync(await saveDocx(parsed, []))
    expect(zip.file('word/media/image1.png')).toBeNull()
    expect(zip.file('word/embeddings/oleObject1.bin')).toBeNull()
    expect(zip.file('word/diagrams/data1.xml')).toBeNull()
    expect(zip.file('customXml/custom-chart.xml')).not.toBeNull()
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).not.toContain('/word/diagrams/data1.xml')
    expect(contentTypes).toContain('/customXml/custom-chart.xml')
  })
})

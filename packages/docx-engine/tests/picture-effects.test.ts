import { describe, expect, it } from 'vitest'
import {
  patchDrawingDocPr,
  patchImageParagraphXml,
  patchShapeStyles,
  patchTableAltText,
  parseDocx,
  saveDocx,
  type ShadowEffect,
} from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML, TINY_PNG_BASE64 } from './helpers/build-docx'

/** inline picture with a full pic:spPr (xfrm + geometry + outline + effects) */
const PIC_SPPR_XML =
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="457200"/>' +
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  '<wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const OFFSET_DIAG_SHADOW: ShadowEffect = {
  blurRadEmu: 50800,
  distEmu: 38100,
  dirEmu: 2700000,
  color: '000000',
  alphaPct: 43,
}

const TEXTBOX_PARAGRAPH =
  '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="1800000" cy="1080000"/>' +
  '<wp:docPr id="2" name="TextBox 2"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="1080000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Shape</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

const TABLE_XML =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableNormal"/><w:tblW w:w="9360" w:type="dxa"/>' +
  '<w:tblLook w:val="04A0"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p/></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
  '<w:p/>'

describe('picture alt text (wp:docPr title/descr)', () => {
  it('adds title and descr with XML escaping', () => {
    const out = patchImageParagraphXml(PIC_SPPR_XML, {
      altTitle: 'A "quote" & <tag>',
      altDescr: 'Logo <1>',
    })
    expect(out).toContain(
      '<wp:docPr id="1" name="Picture 1" title="A &quot;quote&quot; &amp; &lt;tag&gt;" descr="Logo &lt;1&gt;"/>',
    )
  })

  it('replaces and removes existing attributes', () => {
    const withAlt = patchImageParagraphXml(PIC_SPPR_XML, { altTitle: 'Old', altDescr: 'Old text' })
    const replaced = patchImageParagraphXml(withAlt, { altTitle: 'New' })
    expect(replaced).toContain('title="New"')
    expect(replaced).toContain('descr="Old text"')
    const cleared = patchImageParagraphXml(replaced, { altTitle: null, altDescr: null })
    expect(cleared).not.toContain('title=')
    expect(cleared).not.toContain('descr=')
  })

  it('round-trips through the parser', async () => {
    const patched = patchImageParagraphXml(PIC_SPPR_XML, {
      altTitle: 'Chart of Q3 & Q4',
      altDescr: 'Revenue by quarter',
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: patched, withImage: true }))
    expect(doc.blocks[0].imageAltTitle).toBe('Chart of Q3 & Q4')
    expect(doc.blocks[0].imageAltText).toBe('Revenue by quarter')
  })

  it('carries authored alt text onto a newly embedded image', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'image',
        image: {
          base64: TINY_PNG_BASE64,
          mime: 'image/png',
          widthPx: 100,
          heightPx: 80,
          altTitle: 'Inserted logo',
          altText: 'Blue square logo',
        },
      },
    ])
    const doc = await parseDocx(saved)
    const img = doc.blocks.find((b) => b.type === 'image')!
    expect(img.imageAltTitle).toBe('Inserted logo')
    expect(img.imageAltText).toBe('Blue square logo')
  })
})

describe('picture shadow (pic:spPr a:effectLst)', () => {
  it('inserts outerShdw after the outline and pads wp:effectExtent', () => {
    const out = patchImageParagraphXml(PIC_SPPR_XML, { shadow: OFFSET_DIAG_SHADOW })
    expect(out).toContain(
      '<a:outerShdw blurRad="50800" dist="38100" dir="2700000" rotWithShape="0">' +
        '<a:srgbClr val="000000"><a:alpha val="43000"/></a:srgbClr></a:outerShdw>',
    )
    // 45°: dist projection ≈ 26941 EMU, so l/t = blur-dx, r/b = blur+dx
    expect(out).toMatch(/<wp:effectExtent l="23860" t="23860" r="77741" b="77741"\/>/)
  })

  it('places a new outline before an existing effectLst', () => {
    const shadowed = patchImageParagraphXml(PIC_SPPR_XML, { shadow: OFFSET_DIAG_SHADOW })
    const bordered = patchImageParagraphXml(shadowed, {
      border: { color: 'FF0000', widthPt: 2.25 },
    })
    const lnAt = bordered.indexOf('<a:ln w="28575">')
    const effectAt = bordered.indexOf('<a:effectLst>')
    expect(lnAt).toBeGreaterThan(-1)
    expect(lnAt).toBeLessThan(effectAt)
  })

  it('removes the effectLst with shadow: null', () => {
    const shadowed = patchImageParagraphXml(PIC_SPPR_XML, { shadow: OFFSET_DIAG_SHADOW })
    expect(shadowed).toContain('<a:effectLst>')
    const cleared = patchImageParagraphXml(shadowed, { shadow: null })
    expect(cleared).not.toContain('<a:effectLst>')
  })

  it('round-trips outer and inner shadows through the parser', async () => {
    const outer = patchImageParagraphXml(PIC_SPPR_XML, { shadow: OFFSET_DIAG_SHADOW })
    const inner = patchImageParagraphXml(PIC_SPPR_XML, {
      shadow: { ...OFFSET_DIAG_SHADOW, inner: true },
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: outer + inner, withImage: true }))
    expect(doc.blocks[0].imageShadow).toEqual(OFFSET_DIAG_SHADOW)
    expect(doc.blocks[1].imageShadow).toEqual({ ...OFFSET_DIAG_SHADOW, inner: true })
  })
})

describe('shadow merge preserves authored non-shadow effects (BUG-1007)', () => {
  /** Word-authored effectLst: glow + reflection + softEdge, no shadow */
  const GLOW_REFLECTION_SPPR = PIC_SPPR_XML.replace(
    '</pic:spPr>',
    '<a:effectLst><a:glow rad="50800"><a:srgbClr val="4472C4"/></a:glow>' +
      '<a:reflection blurRad="6350" stA="44000" stPos="0" endA="0" endPos="45000" dist="0" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>' +
      '<a:softEdge rad="6350"/></a:effectLst></pic:spPr>',
  )

  it('authoring a shadow keeps glow/reflection/softEdge in the effectLst', () => {
    const out = patchImageParagraphXml(GLOW_REFLECTION_SPPR, { shadow: OFFSET_DIAG_SHADOW })
    expect(out).toContain('<a:glow rad="50800">')
    expect(out).toContain('<a:reflection ')
    expect(out).toContain('<a:softEdge rad="6350"/>')
    expect(out).toContain('<a:outerShdw blurRad="50800" dist="38100" dir="2700000"')
  })

  it('clearing the shadow keeps the sibling effects', () => {
    const shadowed = patchImageParagraphXml(GLOW_REFLECTION_SPPR, { shadow: OFFSET_DIAG_SHADOW })
    const cleared = patchImageParagraphXml(shadowed, { shadow: null })
    expect(cleared).not.toContain('outerShdw')
    expect(cleared).toContain('<a:glow rad="50800">')
    expect(cleared).toContain('<a:softEdge rad="6350"/>')
    expect(cleared).toContain('<a:effectLst>')
  })

  it('replaces an existing Word-authored shadow, keeping the rest', () => {
    const authored = PIC_SPPR_XML.replace(
      '</pic:spPr>',
      '<a:effectLst>' +
        '<a:outerShdw blurRad="12700" dist="0" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"/></a:outerShdw>' +
        '<a:softEdge rad="6350"/></a:effectLst></pic:spPr>',
    )
    const out = patchImageParagraphXml(authored, { shadow: OFFSET_DIAG_SHADOW })
    expect(out.match(/<a:outerShdw/g)).toHaveLength(1)
    expect(out).toContain('blurRad="50800"')
    expect(out).toContain('<a:softEdge rad="6350"/>')
  })

  it('shape styles merge keeps textbox effects too (patchShapeStyles)', () => {
    const withEffects = TEXTBOX_PARAGRAPH.replace(
      '</wps:spPr>',
      '<a:effectLst><a:glow rad="50800"><a:srgbClr val="4472C4"/></a:glow></a:effectLst></wps:spPr>',
    )
    const shadowed = patchShapeStyles(withEffects, [{ shadow: OFFSET_DIAG_SHADOW }])
    expect(shadowed).toContain('<a:outerShdw')
    const cleared = patchShapeStyles(shadowed, [{ shadow: null }])
    expect(cleared).toContain('<a:glow rad="50800">')
    expect(cleared).not.toContain('outerShdw')
  })

  it('a self-closing effectLst still takes the shadow', () => {
    const empty = PIC_SPPR_XML.replace('</pic:spPr>', '<a:effectLst/></pic:spPr>')
    const out = patchImageParagraphXml(empty, { shadow: OFFSET_DIAG_SHADOW })
    expect(out).toContain('<a:effectLst><a:outerShdw')
  })
})

describe('picture outline (pic:spPr a:ln)', () => {
  it('writes the outline with the weight in EMU', () => {
    const out = patchImageParagraphXml(PIC_SPPR_XML, {
      border: { color: '2B5797', widthPt: 1.5 },
    })
    expect(out).toContain(
      '<a:ln w="19050"><a:solidFill><a:srgbClr val="2B5797"/></a:solidFill></a:ln>',
    )
  })

  it('replaces an existing outline and writes explicit noFill on null', () => {
    const bordered = patchImageParagraphXml(PIC_SPPR_XML, {
      border: { color: '2B5797', widthPt: 1.5 },
    })
    const recolored = patchImageParagraphXml(bordered, {
      border: { color: 'C00000', widthPt: 3 },
    })
    expect(recolored).toContain(
      '<a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>',
    )
    expect(recolored).not.toContain('2B5797')
    const cleared = patchImageParagraphXml(recolored, { border: null })
    expect(cleared).toContain('<a:ln><a:noFill/></a:ln>')
  })

  it('round-trips through the parser', async () => {
    const patched = patchImageParagraphXml(PIC_SPPR_XML, {
      border: { color: 'FFD428', widthPt: 2.25 },
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: patched, withImage: true }))
    expect(doc.blocks[0].imageBorder).toEqual({ color: 'FFD428', widthPt: 2.25 })
  })

  it('embeds authored outline and shadow on a new image', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(bytes)
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'image',
        image: {
          base64: TINY_PNG_BASE64,
          mime: 'image/png',
          widthPx: 100,
          heightPx: 80,
          shadow: OFFSET_DIAG_SHADOW,
          border: { color: '2B5797', widthPt: 1.5 },
        },
      },
    ])
    const doc = await parseDocx(saved)
    const img = doc.blocks.find((b) => b.type === 'image')!
    expect(img.imageShadow).toEqual(OFFSET_DIAG_SHADOW)
    expect(img.imageBorder).toEqual({ color: '2B5797', widthPt: 1.5 })
  })
})

describe('shape shadow + alt text (wps:spPr / wp:docPr)', () => {
  it('writes and clears the shape effectLst via patchShapeStyles', () => {
    const shadowed = patchShapeStyles(TEXTBOX_PARAGRAPH, [{ shadow: OFFSET_DIAG_SHADOW }])
    expect(shadowed).toContain('<a:effectLst><a:outerShdw')
    const cleared = patchShapeStyles(shadowed, [{ shadow: null }])
    expect(cleared).not.toContain('<a:effectLst>')
  })

  it('round-trips a shape shadow through the parser', async () => {
    const shadowed = patchShapeStyles(TEXTBOX_PARAGRAPH, [{ shadow: OFFSET_DIAG_SHADOW }])
    const doc = await parseDocx(await buildDocx({ bodyXml: shadowed }))
    expect(doc.blocks[0].textboxes?.[0]?.shadow).toEqual(OFFSET_DIAG_SHADOW)
  })

  it('parses alt text of the shape drawing and patches its docPr', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TEXTBOX_PARAGRAPH }))
    expect(doc.blocks[0].imageAltTitle).toBeUndefined()
    const withAlt = patchDrawingDocPr(TEXTBOX_PARAGRAPH, {
      title: 'Callout',
      descr: 'Explains the chart',
    })
    expect(withAlt).toContain('descr="Explains the chart"')
    const doc2 = await parseDocx(await buildDocx({ bodyXml: withAlt }))
    expect(doc2.blocks[0].imageAltTitle).toBe('Callout')
    expect(doc2.blocks[0].imageAltText).toBe('Explains the chart')
    const cleared = patchDrawingDocPr(withAlt, { title: null, descr: null })
    expect(cleared).not.toContain('descr=')
  })
})

describe('table alt text (w:tblCaption / w:tblDescription)', () => {
  it('adds, replaces and removes both fields', () => {
    const withAlt = patchTableAltText(TABLE_XML, { title: 'Sales', descr: 'Quarterly sales' })
    expect(withAlt).toContain('<w:tblCaption w:val="Sales"/>')
    expect(withAlt).toContain('<w:tblDescription w:val="Quarterly sales"/>')
    const replaced = patchTableAltText(withAlt, { title: 'Revenue' })
    expect(replaced).toContain('<w:tblCaption w:val="Revenue"/>')
    const cleared = patchTableAltText(replaced, { title: null, descr: null })
    expect(cleared).not.toContain('w:tblCaption')
    expect(cleared).not.toContain('w:tblDescription')
  })

  it('round-trips through the parser', async () => {
    const withAlt = patchTableAltText(TABLE_XML, { title: 'Sales', descr: 'Quarterly sales' })
    const doc = await parseDocx(await buildDocx({ bodyXml: withAlt }))
    expect(doc.blocks[0].table?.altTitle).toBe('Sales')
    expect(doc.blocks[0].table?.altText).toBe('Quarterly sales')
  })

  it('escapes attribute values', () => {
    const out = patchTableAltText(TABLE_XML, { title: 'A & B' })
    expect(out).toContain('<w:tblCaption w:val="A &amp; B"/>')
  })
})

describe('image paragraph XML stays untouched without effect patches', () => {
  it('keeps bytes identical when nothing is patched', () => {
    expect(patchImageParagraphXml(IMAGE_PARAGRAPH_XML, {})).toBe(IMAGE_PARAGRAPH_XML)
  })
})

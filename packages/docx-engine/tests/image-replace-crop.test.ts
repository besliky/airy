/**
 * BUG-1006: replacing an original image's bytes (Compress Pictures) without
 * "delete cropped areas" must keep the Word-authored crop window. The replace
 * pipeline used to strip <a:srcRect> and reset <a:fillRect> unconditionally,
 * so the file "uncropped" itself after reopen while the editor still showed
 * the crop. bake flows (crop / cutout / replace) keep stripping it.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

/** inline picture with a Word-authored crop and a non-default fill window */
const CROPPED_IMAGE_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/>' +
  '<a:srcRect l="25000" t="0" r="25000" b="0"/>' +
  '<a:stretch><a:fillRect l="10000" t="0" r="0" b="10000"/></a:stretch></pic:blipFill>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

async function savedDocumentXml(replaceImage: {
  base64: string
  mime: 'image/png'
  keepCrop?: boolean
}): Promise<string> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: CROPPED_IMAGE_XML, withImage: true }))
  const image = parsed.blocks.find((block) => block.type === 'image')!
  const bytes = await saveDocx(parsed, [
    { kind: 'xml', xml: image.originalXml!, docxIndex: image.docxIndex!, replaceImage },
  ])
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

describe('image replace vs crop window (BUG-1006)', () => {
  it('keepCrop keeps the srcRect/fillRect window through the swap', async () => {
    const xml = await savedDocumentXml({
      base64: TINY_PNG_BASE64,
      mime: 'image/png',
      keepCrop: true,
    })
    expect(xml).toContain('<a:srcRect l="25000" t="0" r="25000" b="0"/>')
    expect(xml).toContain('<a:fillRect l="10000" t="0" r="0" b="10000"/>')
  })

  it('bake flows (no keepCrop) still strip the crop and reset the fill window', async () => {
    const xml = await savedDocumentXml({ base64: TINY_PNG_BASE64, mime: 'image/png' })
    expect(xml).not.toContain('<a:srcRect')
    expect(xml).toContain('<a:fillRect/>')
    expect(xml).not.toContain('<a:fillRect l=')
  })
})

/**
 * BUG-1006 (renderer wiring): Compress Pictures without "delete cropped
 * areas" sets imageReplace { keepCrop: true } on the original-image node; the
 * save plan must carry the flag so the engine keeps the Word-authored
 * a:srcRect window in the saved file.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'

const CROPPED_IMAGE_XML =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/>' +
  '<a:srcRect l="25000" t="0" r="25000" b="0"/>' +
  '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const REPLACEMENT_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

async function openCroppedDoc() {
  return parseDocx(await buildDocx({ bodyXml: CROPPED_IMAGE_XML, withImage: true }))
}

describe('compress keep-crop reaches the saved file (BUG-1006)', () => {
  it('imageReplace keepCrop survives the save plan and the save', async () => {
    const parsed = await openCroppedDoc()
    const image = parsed.blocks.find((b) => b.type === 'image')!
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'docProtected',
          attrs: {
            docxIndex: image.docxIndex,
            blockType: 'image',
            label: 'Image',
            imageWidthPx: 96,
            imageHeightPx: 96,
            imageReplace: { base64: REPLACEMENT_PNG, mime: 'image/png', keepCrop: true },
          },
        },
      ],
    } as never as PmNode
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml') as {
      replaceImage?: { keepCrop?: boolean }
    } & { xml: string }
    expect(xmlBlock).toBeDefined()
    expect(xmlBlock.replaceImage?.keepCrop).toBe(true)

    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const zip = await JSZip.loadAsync(bytes)
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).toContain('<a:srcRect l="25000" t="0" r="25000" b="0"/>')
  })

  it('imageReplace without keepCrop still strips the crop (bake flows)', async () => {
    const parsed = await openCroppedDoc()
    const image = parsed.blocks.find((b) => b.type === 'image')!
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'docProtected',
          attrs: {
            docxIndex: image.docxIndex,
            blockType: 'image',
            label: 'Image',
            imageWidthPx: 96,
            imageHeightPx: 96,
            imageReplace: { base64: REPLACEMENT_PNG, mime: 'image/png' },
          },
        },
      ],
    } as never as PmNode
    const plan = pmDocToSavePlan(doc, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const zip = await JSZip.loadAsync(bytes)
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).not.toContain('<a:srcRect')
  })
})

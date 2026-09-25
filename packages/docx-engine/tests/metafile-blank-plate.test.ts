import { describe, expect, it } from 'vitest'
import { buildDocx } from './helpers/build-docx'
import { metafilePlateText, metafileToDataUrl } from '../src/metafile'
import { parseDocx } from '../src/index'

/**
 * UX-1710: metafile pictures that cannot be shown (conversion failed, or — the
 * live-canvas case — a valid header with no drawable records rendering a fully
 * transparent frame) must degrade to the broken-image plate naming the part
 * kind and size ("WMF image (1 KB)"), not a silent empty rectangle.
 */

/** placeable header + WMF header + EOF record: minimal well-formed WMF */
function minimalWmf(): Uint8Array {
  const u8 = new Uint8Array(46)
  const dv = new DataView(u8.buffer)
  dv.setUint32(0, 0x9ac6cdd7, true) // placeable magic
  dv.setUint16(4, 0, true) // hwmf
  dv.setInt16(6, 0, true) // bboxLeft
  dv.setInt16(8, 0, true) // bboxTop
  dv.setInt16(10, 1000, true) // bboxRight
  dv.setInt16(12, 1000, true) // bboxBottom
  dv.setUint16(14, 1440, true) // inch
  dv.setUint16(20, 0, true) // checksum
  dv.setUint16(22, 2, true) // header: disk file
  dv.setUint16(24, 9, true) // headerSize (words)
  dv.setUint16(26, 0x0300, true) // version
  dv.setUint32(28, 23, true) // fileSize (words)
  dv.setUint16(32, 0, true) // numOfObjects
  dv.setUint32(34, 3, true) // maxRecord (words)
  dv.setUint16(38, 0, true) // numOfParams
  dv.setUint32(40, 3, true) // EOF record: size (words)
  dv.setUint16(44, 0, true) // EOF record: function
  return u8
}

/** EMR_HEADER + EMR_EOF: minimal well-formed EMF (108 bytes, the audit's shape) */
function minimalEmf(): Uint8Array {
  const u8 = new Uint8Array(108)
  const dv = new DataView(u8.buffer)
  dv.setUint32(0, 1, true) // EMR_HEADER
  dv.setUint32(4, 88, true) // nSize
  dv.setInt32(8, 0, true) // bounds left
  dv.setInt32(12, 0, true) // bounds top
  dv.setInt32(16, 1000, true) // bounds right
  dv.setInt32(20, 1000, true) // bounds bottom
  dv.setInt32(24, 0, true) // frame left
  dv.setInt32(28, 0, true) // frame top
  dv.setInt32(32, 9525, true) // frame right (0.01mm)
  dv.setInt32(36, 9525, true) // frame bottom
  dv.setUint32(40, 0x464d4520, true) // ' EMF' signature
  dv.setUint32(44, 0x00010000, true) // version
  dv.setUint32(48, 108, true) // total bytes
  dv.setUint32(52, 2, true) // record count
  dv.setUint16(56, 1, true) // handle count
  dv.setUint16(58, 0, true) // reserved
  dv.setUint32(60, 0, true) // nDescription
  dv.setUint32(64, 0, true) // offDescription
  dv.setUint32(68, 0, true) // nPalEntries
  dv.setInt32(72, 1920, true) // szlDevice cx
  dv.setInt32(76, 1080, true) // szlDevice cy
  dv.setInt32(80, 508, true) // szlMillimeters cx
  dv.setInt32(84, 285, true) // szlMillimeters cy
  dv.setUint32(88, 14, true) // EMR_EOF
  dv.setUint32(92, 20, true) // nSize
  dv.setUint32(96, 0, true) // nPalEntries
  dv.setUint32(100, 16, true) // offPalEntries
  dv.setUint32(104, 20, true) // nSizeLast
  return u8
}

const REL = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`

function drawingXml(rid: string): string {
  return (
    '<w:drawing><wp:inline><wp:extent cx="1828800" cy="1828800"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill></pic:pic>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing>'
  )
}

async function parseWithMetafile(
  rid: string,
  target: string,
  base64: string,
  extension: string,
  contentType: string,
) {
  const source = await buildDocx({
    bodyXml: `<w:p><w:r>${drawingXml(rid)}</w:r></w:p>`,
    extraRels: REL(rid, target),
    binaryParts: [{ path: `word/${target}`, base64, extension, contentType }],
  })
  return parseDocx(source)
}

describe('UX-1710: unshowable metafile pictures name the part on the plate', () => {
  it('metafilePlateText formats kind and size', () => {
    expect(metafilePlateText('image/wmf', 36)).toBe('WMF image (1 KB)')
    expect(metafilePlateText('image/x-emf', 2048)).toBe('EMF image (2 KB)')
    expect(metafilePlateText('image/x-wmz', 4096)).toBe('WMZ image (4 KB)')
  })

  it('blank-or-failed WMF degrades to a broken-image plate naming the part', async () => {
    // no canvas API under vitest, so the converter itself returns null — the
    // same degrade the live renderer now reaches for blank renders
    expect(await metafileToDataUrl(minimalWmf(), 'image/wmf')).toBeNull()
    const parsed = await parseWithMetafile(
      'rIdWmf',
      'media/image1.wmf',
      Buffer.from(minimalWmf()).toString('base64'),
      'wmf',
      'image/wmf',
    )
    const block = parsed.blocks.find((b) => b.brokenImage)
    expect(block).toBeDefined()
    expect(block?.previewText).toBe('WMF image (1 KB)')
  })

  it('blank-or-failed EMF degrades the same way', async () => {
    const parsed = await parseWithMetafile(
      'rIdEmf',
      'media/image2.emf',
      Buffer.from(minimalEmf()).toString('base64'),
      'emf',
      'image/emf',
    )
    const block = parsed.blocks.find((b) => b.brokenImage)
    expect(block).toBeDefined()
    expect(block?.previewText).toBe('EMF image (1 KB)')
  })

  it('an author alt text still wins over the part label', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Pic" descr="Legacy logo"/>' +
        '<wp:extent cx="1828800" cy="1828800"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic><pic:blipFill><a:blip r:embed="rIdWmf"/></pic:blipFill></pic:pic>' +
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
      extraRels: REL('rIdWmf', 'media/image1.wmf'),
      binaryParts: [
        {
          path: 'word/media/image1.wmf',
          base64: Buffer.from(minimalWmf()).toString('base64'),
          extension: 'wmf',
          contentType: 'image/wmf',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const block = parsed.blocks.find((b) => b.brokenImage)
    expect(block?.previewText).toBe('Legacy logo')
  })
})

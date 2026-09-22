import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { xmlWellFormednessError } from '../src/xml-utils'
import { buildDocx } from './helpers/build-docx'

const PLAIN_PARA = '<w:p><w:r><w:t>hello</w:t></w:r></w:p>'
const SECOND_PARA = '<w:p><w:r><w:t>second</w:t></w:r></w:p>'

/** the document.xml that buildDocx writes for the two-paragraph body */
async function baseDocumentXml(): Promise<string> {
  const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: PLAIN_PARA + SECOND_PARA }))
  return zip.file('word/document.xml')!.async('string')
}

/** rebuild the package with word/document.xml replaced by the given raw XML */
async function withDocumentXml(documentXml: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: PLAIN_PARA }))
  zip.file('word/document.xml', documentXml)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** the base document with the body content swapped (root tag and namespaces kept) */
async function documentXmlWithBody(bodyXml: string): Promise<string> {
  const full = await baseDocumentXml()
  const start = full.indexOf('<w:body>') + '<w:body>'.length
  const end = full.indexOf('</w:body>')
  return full.slice(0, start) + bodyXml + full.slice(end)
}

describe('BUG-1609: a corrupt docx refuses loudly instead of opening as an empty document', () => {
  it('refuses a document.xml cut mid-tag (the audit patho case)', async () => {
    const full = await baseDocumentXml()
    const secondAt = full.indexOf(SECOND_PARA)
    const bytes = await withDocumentXml(full.slice(0, secondAt + 10)) // "<w:p><w:r><"
    await expect(parseDocx(bytes)).rejects.toThrow(
      /docx file is corrupted: word\/document\.xml is not well-formed XML/,
    )
  })

  it('refuses a document.xml cut mid-attribute', async () => {
    const bytes = await withDocumentXml(
      await documentXmlWithBody(PLAIN_PARA + '<w:p w:rsidR="00A1'),
    )
    await expect(parseDocx(bytes)).rejects.toThrow(/not well-formed XML/)
  })

  it('refuses a document.xml cut right after <w:body>', async () => {
    const full = await baseDocumentXml()
    const bytes = await withDocumentXml(full.slice(0, full.indexOf('<w:body>') + 8))
    await expect(parseDocx(bytes)).rejects.toThrow(/not well-formed XML/)
  })

  it('refuses a document.xml cut inside the closing tag', async () => {
    const full = await baseDocumentXml()
    const bytes = await withDocumentXml(full.slice(0, full.length - 5)) // "</w:documen"
    await expect(parseDocx(bytes)).rejects.toThrow(/not well-formed XML/)
  })

  it('refuses mismatched closing tags in the body', async () => {
    const bytes = await withDocumentXml(await documentXmlWithBody('<w:p><w:r><w:t>x</w:p></w:r>'))
    await expect(parseDocx(bytes)).rejects.toThrow(/not well-formed XML/)
  })

  it('reports the validator position in the refusal', async () => {
    const full = await baseDocumentXml()
    const bytes = await withDocumentXml(full.slice(0, full.length - 5))
    await expect(parseDocx(bytes)).rejects.toThrow(/\(line \d+, col \d+\)/)
  })

  it('refuses a zip truncated before the central directory', async () => {
    const bytes = await buildDocx({ bodyXml: PLAIN_PARA })
    await expect(parseDocx(bytes.slice(0, bytes.length - 40))).rejects.toThrow(
      /docx file is corrupted: not a readable zip package/,
    )
  })

  it('refuses bytes that are not a zip archive at all', async () => {
    const bytes = new TextEncoder().encode('plainly not a zip archive, just text')
    await expect(parseDocx(bytes)).rejects.toThrow(/not a readable zip package/)
  })

  it('refuses a package whose main part is missing', async () => {
    const zip = await JSZip.loadAsync(await buildDocx({ bodyXml: PLAIN_PARA }))
    zip.remove('word/document.xml')
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
    await expect(parseDocx(bytes)).rejects.toThrow(/not a docx: missing word\/document\.xml/)
  })
})

describe('BUG-1609: files with recoverable quirks still open (no false refusals)', () => {
  it('opens an intact two-paragraph document', async () => {
    const doc = await parseDocx(await withDocumentXml(await baseDocumentXml()))
    // buildDocx bodies carry a trailing body-level sectPr, surfaced as a hidden
    // passthrough block after the two paragraphs
    expect(doc.blocks).toHaveLength(3)
    expect(doc.blocks[0].runs?.[0]?.text).toBe('hello')
    expect(doc.blocks[1].runs?.[0]?.text).toBe('second')
  })

  it('opens a document carrying an undefined entity reference as literal text', async () => {
    const bytes = await withDocumentXml(
      await documentXmlWithBody('<w:p><w:r><w:t>&lol9; laughs</w:t></w:r></w:p>'),
    )
    const doc = await parseDocx(bytes)
    expect(JSON.stringify(doc.blocks)).toContain('lol9')
  })

  it('opens deeply nested (3000 levels) markup, degrading the block instead of refusing', async () => {
    const DEEP = 3000
    const bytes = await withDocumentXml(
      await documentXmlWithBody(
        '<w:p>' +
          '<w:smartTag>'.repeat(DEEP) +
          '<w:r><w:t>deep</w:t></w:r>' +
          '</w:smartTag>'.repeat(DEEP) +
          '</w:p>' +
          PLAIN_PARA,
      ),
    )
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].type).toBe('passthrough')
    expect(doc.blocks[0].previewText).toContain('deep')
    expect(doc.blocks[1].runs?.[0]?.text).toBe('hello')
  })

  it('treats XML comments and CDATA in the main part as well-formed (gate stays silent)', async () => {
    // scanBody cannot walk a commented body (pre-existing, throws its own
    // error); the BUG-1609 gate must not add a second refusal on top
    const xml = await documentXmlWithBody(
      '<w:p><!-- parser note --><w:r><w:t><![CDATA[x < y]]> ok</w:t></w:r></w:p>',
    )
    expect(xmlWellFormednessError(xml)).toBeNull()
  })

  it('opens Word-tolerated sibling w:body elements (POI MultipleBodyBug shape)', async () => {
    const full = await baseDocumentXml()
    const start = full.indexOf('<w:body>') + '<w:body>'.length
    const end = full.indexOf('</w:body>')
    const inner = full.slice(start, end)
    const twoBodies =
      full.slice(0, start) + SECOND_PARA + '</w:body><w:body>' + inner + full.slice(end)
    const doc = await parseDocx(await withDocumentXml(twoBodies))
    const allText = JSON.stringify(doc.blocks)
    expect(allText).toContain('second')
    expect(allText).toContain('hello')
  })
})

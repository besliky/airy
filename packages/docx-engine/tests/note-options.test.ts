import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const SETTINGS_REL =
  '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'

const SETTINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'

function settingsDocx(settingsInner: string, body = '<w:p><w:r><w:t>a</w:t></w:r></w:p>') {
  return buildDocx({
    bodyXml: body,
    extraRels: SETTINGS_REL,
    extraParts: [
      {
        path: 'word/settings.xml',
        contentType: SETTINGS_CONTENT_TYPE,
        xml:
          XML_DECL +
          '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          settingsInner +
          '</w:settings>',
      },
    ],
  })
}

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

describe('settings.xml w:footnotePr / w:endnotePr parse', () => {
  it('reads numFmt / numStart / numRestart of both kinds', async () => {
    const doc = await parseDocx(
      await settingsDocx(
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="2"/>' +
          '<w:numRestart w:val="eachPage"/></w:footnotePr>' +
          '<w:endnotePr><w:numFmt w:val="upperRoman"/></w:endnotePr>',
      ),
    )
    expect(doc.noteNumbering).toEqual({
      footnotes: { numFmt: 'lowerLetter', numStart: 2, numRestart: 'eachPage' },
      endnotes: { numFmt: 'upperRoman' },
    })
  })

  it('absent tags = undefined; unmodeled children alone stay undefined', async () => {
    expect((await parseDocx(await settingsDocx(''))).noteNumbering).toBeUndefined()
    // w:pos is not modeled: the tag pair must not surface as empty defaults
    expect(
      (await parseDocx(await settingsDocx('<w:footnotePr><w:pos w:val="pageBottom"/></w:footnotePr>')))
        .noteNumbering,
    ).toBeUndefined()
  })
})

describe('SaveOptions.noteNumbering (Word note-options dialog)', () => {
  it('writes both tags into an existing settings.xml and round-trips', async () => {
    const doc = await parseDocx(await settingsDocx(''))
    const out = await saveDocx(doc, originalOrder(doc), {
      noteNumbering: {
        footnotes: { numFmt: 'lowerRoman', numStart: 3, numRestart: 'eachSect' },
        endnotes: { numFmt: 'upperLetter' },
      },
    })
    const settingsXml = await (await JSZip.loadAsync(out)).file('word/settings.xml')!.async('string')
    expect(settingsXml).toContain(
      '<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/>' +
        '<w:numRestart w:val="eachSect"/></w:footnotePr>',
    )
    expect(settingsXml).toContain('<w:endnotePr><w:numFmt w:val="upperLetter"/></w:endnotePr>')
    // schema position: footnotePr/endnotePr land before w:compat when present
    const compat = '<w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat>'
    const doc2 = await parseDocx(await settingsDocx(compat))
    const out2 = await saveDocx(doc2, originalOrder(doc2), {
      noteNumbering: { footnotes: { numFmt: 'decimal' } },
    })
    const xml2 = await (await JSZip.loadAsync(out2)).file('word/settings.xml')!.async('string')
    expect(xml2.indexOf('<w:footnotePr>')).toBeLessThan(xml2.indexOf('<w:compat'))
    // full round-trip through the parser
    const reparsed = await parseDocx(out)
    expect(reparsed.noteNumbering).toEqual({
      footnotes: { numFmt: 'lowerRoman', numStart: 3, numRestart: 'eachSect' },
      endnotes: { numFmt: 'upperLetter' },
    })
  })

  it('creates settings.xml (with rel + content type) when the package has none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('a') }))
    const out = await saveDocx(doc, originalOrder(doc), {
      noteNumbering: { endnotes: { numFmt: 'lowerRoman' } },
    })
    const zip = await JSZip.loadAsync(out)
    expect(zip.file('word/settings.xml')).not.toBeNull()
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('relationships/settings"')
    expect(rels).toContain('Target="settings.xml"')
    const reparsed = await parseDocx(out)
    expect(reparsed.noteNumbering).toEqual({ endnotes: { numFmt: 'lowerRoman' } })
  })

  it('null removes the tag; undefined keeps it untouched', async () => {
    const doc = await parseDocx(
      await settingsDocx(
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/></w:footnotePr>' +
          '<w:endnotePr><w:numFmt w:val="upperRoman"/></w:endnotePr>',
      ),
    )
    const out = await saveDocx(doc, originalOrder(doc), {
      noteNumbering: { footnotes: null, endnotes: undefined },
    })
    const settingsXml = await (await JSZip.loadAsync(out)).file('word/settings.xml')!.async('string')
    expect(settingsXml).not.toContain('footnotePr')
    expect(settingsXml).toContain('<w:endnotePr><w:numFmt w:val="upperRoman"/></w:endnotePr>')
    // no option at all: bytes stay identical
    const untouched = await saveDocx(doc, originalOrder(doc))
    expect(untouched).toBe(doc.internal.originalBytes)
  })
})

describe('custom note marks (w:customMarkFollows)', () => {
  const FOOTNOTES_XML =
    XML_DECL +
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note</w:t></w:r></w:p></w:footnote>' +
    '</w:footnotes>'

  const FOOTNOTES_REL =
    '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'

  const CUSTOM_MARK_P =
    '<w:p><w:r><w:t>a</w:t></w:r>' +
    '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' +
    '<w:footnoteReference w:id="2" w:customMarkFollows="1"/><w:t>*</w:t></w:r></w:p>'

  it('parses the literal mark and regenerates customMarkFollows on save', async () => {
    const bytes = await buildDocx({
      bodyXml: CUSTOM_MARK_P,
      extraRels: FOOTNOTES_REL,
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml: FOOTNOTES_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs).toEqual([
      { text: 'a' },
      { text: '*', noteRef: { kind: 'footnote', id: '2', customMark: '*' } },
    ])
    // untouched original paragraph: bytes identical (mark lives in the source)
    const out = await saveDocx(doc, originalOrder(doc))
    expect(out).toBe(doc.internal.originalBytes)
    // regenerated paragraph carries customMarkFollows + the literal mark
    const out2 = await saveDocx(doc, [
      ...originalOrder(doc).slice(1),
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'b', noteRef: { kind: 'footnote', id: '2', customMark: '†' } }],
        },
      },
    ])
    const bodyXml = await (
      await JSZip.loadAsync(out2)
    ).file('word/document.xml')!.async('string')
    expect(bodyXml).toContain(
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' +
        '<w:footnoteReference w:id="2" w:customMarkFollows="1"/><w:t xml:space="preserve">†</w:t></w:r>',
    )
  })

  it('a rebuilt note entry prints the mark instead of the self-reference run', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('a') }))
    const out = await saveDocx(
      doc,
      [
        {
          kind: 'generated',
          block: {
            type: 'paragraph',
            runs: [{ text: 'a', noteRef: { kind: 'footnote', id: '2', customMark: '*' } }],
          },
        },
      ],
      { footnotes: [{ id: '2', text: 'custom note', customMark: '*' }] },
    )
    const fnXml = await (await JSZip.loadAsync(out)).file('word/footnotes.xml')!.async('string')
    expect(fnXml).not.toContain('w:footnoteRef')
    expect(fnXml).toContain('<w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">*</w:t>')
    // re-parse: the entry has no self-reference mark → noRefMark, text keeps the mark like Word
    const reparsed = await parseDocx(out)
    expect(reparsed.footnotes).toEqual([{ id: '2', text: '* custom note', noRefMark: true }])
  })
})

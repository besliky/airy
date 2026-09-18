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

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

async function settingsOf(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/settings.xml')!.async('string')
}

describe('w:autoHyphenation / w:hyphenationZone parse', () => {
  it('reads the flag and the zone (twips)', async () => {
    const doc = await parseDocx(
      await settingsDocx('<w:autoHyphenation/><w:hyphenationZone w:val="425"/>'),
    )
    expect(doc.autoHyphenation).toBe(true)
    expect(doc.hyphenationZoneTwips).toBe(425)
  })

  it('absent flag = undefined; w:val="0|false" counts as off', async () => {
    expect((await parseDocx(await settingsDocx(''))).autoHyphenation).toBeUndefined()
    expect(
      (await parseDocx(await settingsDocx('<w:autoHyphenation w:val="false"/>'))).autoHyphenation,
    ).toBeUndefined()
  })
})

describe('SaveOptions.hyphenation (Layout → Hyphenation authoring)', () => {
  it('auto:true writes the flag (creating settings.xml when absent) and round-trips', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>' }))
    const out = await saveDocx(doc, originalOrder(doc), { hyphenation: { auto: true } })
    const xml = await settingsOf(out)
    expect(xml).toContain('<w:autoHyphenation/>')
    const reparsed = await parseDocx(out)
    expect(reparsed.autoHyphenation).toBe(true)
  })

  it('auto:false removes an existing flag; zone follows the flag', async () => {
    const doc = await parseDocx(
      await settingsDocx('<w:autoHyphenation/><w:hyphenationZone w:val="360"/>'),
    )
    const out = await saveDocx(doc, originalOrder(doc), { hyphenation: { auto: false } })
    const xml = await settingsOf(out)
    expect(xml).not.toContain('autoHyphenation')
    // Word keeps w:hyphenationZone when the flag goes off (it is inert alone);
    // only an explicit zoneTwips:null removes it
    expect(xml).toContain('<w:hyphenationZone w:val="360"/>')
  })

  it('zone rewrite keeps the flag and replaces the value', async () => {
    const doc = await parseDocx(
      await settingsDocx('<w:autoHyphenation/><w:hyphenationZone w:val="360"/>'),
    )
    const out = await saveDocx(doc, originalOrder(doc), {
      hyphenation: { auto: true, zoneTwips: 567 },
    })
    const xml = await settingsOf(out)
    expect(xml).toContain('<w:autoHyphenation/><w:hyphenationZone w:val="567"/>')
    expect((xml.match(/<w:hyphenationZone/g) ?? []).length).toBe(1)
    expect((await parseDocx(out)).hyphenationZoneTwips).toBe(567)
  })

  it('zone:null removes the tag; no hyphenation option keeps bytes identical', async () => {
    const doc = await parseDocx(
      await settingsDocx('<w:autoHyphenation/><w:hyphenationZone w:val="360"/>'),
    )
    const out = await saveDocx(doc, originalOrder(doc), { hyphenation: { zoneTwips: null } })
    expect(await settingsOf(out)).not.toContain('hyphenationZone')
    expect(await settingsOf(out)).toContain('<w:autoHyphenation/>')
    // untouched save: byte-identical package
    const unchanged = await saveDocx(doc, originalOrder(doc))
    expect(unchanged).toBe(doc.internal.originalBytes)
  })
})

describe('soft hyphen round-trip (w:softHyphen)', () => {
  it('a generated paragraph keeps U+00AD as <w:softHyphen/> and re-parses', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>' }))
    const out = await saveDocx(doc, [
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'hy\u00adphen' }] },
      },
    ])
    const bodyXml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(bodyXml).toContain('<w:softHyphen/>')
    const reparsed = await parseDocx(out)
    const para = reparsed.blocks.find((b) => b.type === 'paragraph')
    expect(para && 'runs' in para ? para.runs?.[0]?.text : undefined).toBe('hy\u00adphen')
  })
})

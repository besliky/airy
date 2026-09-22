import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// BUG-1502: a paragraph whose numbering lives in its pStyle (ListBullet with
// w:numPr on the style) must stay unlisted when a regenerated block keeps the
// style: the rebuilt pPr carries Word's explicit "no numbering" override
// (w:numPr with numId="0"), or the saved file resurrects the list both in
// Word and in our own parser.

/** ListBullet/ListNumber styles with numbering on the style, not the paragraph */
const LIST_STYLES =
  '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quiet"><w:name w:val="Quiet"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="120"/></w:pPr></w:style>'

const STYLE_BULLET_P =
  '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Styled bullet</w:t></w:r></w:p>'

async function buildStyleListDocx(bodyXml: string): Promise<Uint8Array> {
  return buildDocx({
    bodyXml,
    extraStylesXml: LIST_STYLES,
    withNumbering: true,
  })
}

describe('style-driven list numbering on regenerated paragraphs (BUG-1502)', () => {
  it('parses a pStyle-numbered paragraph as a listItem without direct numPr', async () => {
    const doc = await parseDocx(await buildStyleListDocx(STYLE_BULLET_P))
    expect(doc.blocks[0].type).toBe('listItem')
    expect(doc.blocks[0].list?.kind).toBe('bullet')
  })

  it('clearList-style regeneration cancels the style numbering with numId="0"', async () => {
    const doc = await parseDocx(await buildStyleListDocx(STYLE_BULLET_P))
    const original = doc.blocks[0]
    // what clearList produces: same style (fonts/indents kept), type paragraph
    const cleared = {
      type: 'paragraph' as const,
      styleId: original.styleId,
      runs: original.runs ?? [],
    }
    const saved = await saveDocx(doc, [{ kind: 'generated', block: cleared }])
    const xml = (await parseDocx(saved)).internal.documentXml
    // Word's own unlisting of a ListBullet paragraph: keep the pStyle, add the
    // explicit no-numbering override after it
    expect(xml).toContain('<w:pStyle w:val="ListBullet"/>')
    expect(xml).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>')
    // and the engine's own parser must not resurrect the list on reopen
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].type).toBe('paragraph')
    expect(reparsed.blocks[0].list).toBeUndefined()
  })

  it('regenerated list items keep their real numPr (no override noise)', async () => {
    const doc = await parseDocx(await buildStyleListDocx(STYLE_BULLET_P))
    const original = doc.blocks[0]
    const asList = {
      type: 'listItem' as const,
      styleId: original.styleId,
      list: original.list,
      runs: original.runs ?? [],
    }
    const saved = await saveDocx(doc, [{ kind: 'generated', block: asList }])
    const xml = (await parseDocx(saved)).internal.documentXml
    expect(xml).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    expect(xml).not.toContain('w:numId w:val="0"')
  })

  it('styles without numbering never gain the override', async () => {
    const body = '<w:p><w:pPr><w:pStyle w:val="Quiet"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildStyleListDocx(body))
    const original = doc.blocks[0]
    expect(original.type).toBe('paragraph')
    const regenerated = {
      type: 'paragraph' as const,
      styleId: original.styleId,
      runs: original.runs ?? [],
    }
    const saved = await saveDocx(doc, [{ kind: 'generated', block: regenerated }])
    const xml = (await parseDocx(saved)).internal.documentXml
    expect(xml).toContain('<w:pStyle w:val="Quiet"/>')
    expect(xml).not.toContain('w:numPr')
  })
})

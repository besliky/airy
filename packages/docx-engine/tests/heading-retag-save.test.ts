// BUG-1501: setHeadingLevel edits were silently reverted on save. The op (and
// the raw pPr passthrough) leave the block's styleId pointing at the OLD
// level's style, and the serializer preferred it over the level mapping, so
// the saved file reopened with the old heading level (or as a heading again
// after a demotion). These tests pin the explicit-level behavior and every
// round-trip case that must keep its bytes.
import { describe, expect, it } from 'vitest'
import {
  generateParagraphXml,
  parseDocx,
  saveDocx,
  type Block,
  type GenerateContext,
  type GeneratedBlock,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

const EXTRA_STYLES =
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="MyHead2"><w:name w:val="My Head 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/></w:style>'

/** audit matrix corpus: h1, plain tail, h2, custom heading, direct-level heading, quoted body, explicit body text */
const BODY = [
  p('<w:pStyle w:val="Heading1"/>', '第一章 概述'),
  p('<w:pStyle w:val="Heading2"/>', '1.1 列表'),
  '<w:p><w:r><w:t>尾段。</w:t></w:r></w:p>',
  p('<w:pStyle w:val="MyHead2"/>', 'custom heading'),
  p('<w:pStyle w:val="Heading2"/><w:outlineLvl w:val="0"/>', 'heading with direct level'),
  p('<w:pStyle w:val="Quote"/>', 'quoted body'),
  p('<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>', 'explicit body text'),
].join('')

const CTX: GenerateContext = {
  headingStyleIds: new Map([
    [1, 'Heading1'],
    [2, 'Heading2'],
    [3, 'Heading3'],
  ]),
  headingLevelOfStyles: new Map<string, number | undefined>([
    ['Heading1', 1],
    ['Heading2', 2],
    ['Heading3', 3],
    ['MyHead2', 2],
    ['Quote', undefined],
    ['Normal', undefined],
  ]),
  allocateHyperlinkRel: () => 'rId1',
}

describe('generateParagraphXml honors an explicit heading level (BUG-1501)', () => {
  it('retag heading→heading rewrites pStyle to the level mapping', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading3"/>')
    expect(xml).not.toContain('Heading2')
  })

  it('retag heading→plain drops the stale heading pStyle', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', styleId: 'Heading1', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).not.toContain('w:pStyle')
  })

  it('plain→heading without a styleId keeps using the level mapping', () => {
    const xml = generateParagraphXml({ type: 'heading', level: 2, runs: [{ text: 'x' }] }, CTX)
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>')
  })

  it('plain→heading over a stale non-heading styleId uses the level mapping too', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 2, styleId: 'Quote', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>')
  })

  it('an unretagged native heading keeps its own styleId (round-trip)', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 2, styleId: 'Heading2', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>')
  })

  it('a custom heading style of the same level keeps its styleId (round-trip)', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 2, styleId: 'MyHead2', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="MyHead2"/>')
  })

  it('a heading with styleId but no explicit level keeps the styleId (old behavior)', () => {
    const xml = generateParagraphXml(
      { type: 'heading', styleId: 'Heading3', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading3"/>')
  })

  it('an outline-only heading keeps its branch untouched by the retag logic', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', outlineOnly: true, runs: [{ text: 'x' }] },
      CTX,
    )
    // pre-fix behavior preserved: the outline level is written, the styleId is
    // left exactly as the model carries it
    expect(xml).toContain('<w:outlineLvl w:val="2"/>')
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>')
  })

  it('a body paragraph keeps a non-heading styleId', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', styleId: 'Quote', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Quote"/>')
  })
})

describe('raw pPr passthrough after a heading retag (BUG-1501)', () => {
  it('swaps the stale pStyle in place, keeping the other pPr bytes', () => {
    const rawPPr =
      '<w:pPr><w:pStyle w:val="Heading2"/><w:keepNext/><w:spacing w:after="120"/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain(
      '<w:pPr><w:pStyle w:val="Heading3"/><w:keepNext/><w:spacing w:after="120"/></w:pPr>',
    )
  })

  it('rewrites a stale direct w:outlineLvl along with the pStyle', () => {
    const rawPPr =
      '<w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="0"/><w:keepNext/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading3"/>')
    expect(xml).toContain('<w:outlineLvl w:val="2"/>')
    expect(xml).toContain('<w:keepNext/>')
  })

  it('keeps the bytes when a direct w:outlineLvl still matches the level', () => {
    const rawPPr =
      '<w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="0"/><w:keepNext/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 1, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain(rawPPr)
  })

  it('keeps a pPrChange revision snapshot untouched while retargeting the outer pPr', () => {
    const rawPPr =
      '<w:pPr><w:pStyle w:val="Heading2"/>' +
      '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">' +
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading3"/>')
    // the OLD-properties snapshot inside the revision keeps its original style
    expect(xml).toContain(
      '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>',
    )
  })

  it('keeps the bytes when the pStyle already agrees with the level', () => {
    const rawPPr = '<w:pPr><w:pStyle w:val="Heading2"/><w:keepNext/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 2, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain(rawPPr)
  })

  it('keeps a custom heading style of the matching level (round-trip)', () => {
    const rawPPr = '<w:pPr><w:pStyle w:val="MyHead2"/><w:keepNext/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 2, styleId: 'MyHead2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain(rawPPr)
  })

  it('leaves non-heading raw pPr alone', () => {
    const rawPPr = '<w:pPr><w:pStyle w:val="Quote"/><w:keepNext/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'paragraph', styleId: 'Quote', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain(rawPPr)
  })
})

// ---- save/repopen round-trips, driving saveDocx the way an editor session
// does: the retag mutates type/level but neither styleId nor the raw pPr
// passthrough (the structure-same contract of the session's toSaveBlocks) ----

/**
 * Minimal port of the editor-session edit contract (packages' editEntry +
 * toSaveBlocks): clone the parsed block, apply the setHeadingLevel mutation,
 * and pass the original raw pPr through when type/styleId/list are unchanged.
 */
function retag(
  block: Block,
  patch: { type?: 'paragraph' | 'heading'; level?: number },
): GeneratedBlock {
  const type = patch.type ?? block.type
  const gen: GeneratedBlock = {
    type: type === 'heading' || type === 'paragraph' ? type : 'paragraph',
    ...(block.level !== undefined || patch.level !== undefined
      ? { level: patch.level ?? block.level }
      : {}),
    ...(block.outlineOnly ? { outlineOnly: true } : {}),
    ...(block.styleId !== undefined ? { styleId: block.styleId } : {}),
    ...(block.format ? (JSON.parse(JSON.stringify(block.format)) as typeof block.format) : {}),
    runs: (block.runs ?? []).map((run) => ({ ...run })),
  }
  if (patch.type === 'paragraph') {
    delete gen.level
    delete gen.outlineOnly
  }
  const structureSame =
    block.type === gen.type &&
    (block.styleId ?? null) === (gen.styleId ?? null) &&
    JSON.stringify(block.list ?? null) === JSON.stringify(gen.list ?? null)
  if (structureSame && block.rawPPr !== undefined) gen.rawPPr = block.rawPPr
  return gen
}

async function fixture() {
  return parseDocx(await buildDocx({ extraStylesXml: EXTRA_STYLES, bodyXml: BODY }))
}

/** final blocks with the given in-place retags applied over the parsed model */
function finalBlocksWith(
  doc: Awaited<ReturnType<typeof fixture>>,
  edits: Map<number, GeneratedBlock>,
): SaveBlock[] {
  return doc.blocks.map((block, i) =>
    edits.has(i)
      ? { kind: 'generated', block: edits.get(i)! }
      : { kind: 'original', docxIndex: block.docxIndex! },
  )
}

describe('heading retag survives save and reopen (BUG-1501)', () => {
  it('heading→heading: the saved file reopens at the new level', async () => {
    const doc = await fixture()
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[1, retag(doc.blocks[1]!, { level: 3 })]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[1]!.type).toBe('heading')
    expect(reopened.blocks[1]!.level).toBe(3)
    expect(reopened.internal.documentXml).toContain('<w:pStyle w:val="Heading3"/>')
  })

  it('heading→plain: the saved file reopens as a body paragraph', async () => {
    const doc = await fixture()
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[0, retag(doc.blocks[0]!, { type: 'paragraph', level: 0 })]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[0]!.type).toBe('paragraph')
    expect(reopened.blocks[0]!.level).toBeUndefined()
    expect(reopened.blocks[0]!.styleId).toBeUndefined()
  })

  it('plain→heading: the saved file reopens at the new level', async () => {
    const doc = await fixture()
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[2, retag(doc.blocks[2]!, { type: 'heading', level: 2 })]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[2]!.type).toBe('heading')
    expect(reopened.blocks[2]!.level).toBe(2)
    expect(reopened.internal.documentXml).toContain('<w:pStyle w:val="Heading2"/>')
  })

  it('saving the same retag twice is byte-identical (idempotent)', async () => {
    const doc = await fixture()
    const edits = new Map([[1, retag(doc.blocks[1]!, { level: 3 })]])
    const first = await saveDocx(doc, finalBlocksWith(doc, edits))
    const second = await saveDocx(doc, finalBlocksWith(doc, edits))
    expect(Buffer.compare(Buffer.from(first), Buffer.from(second))).toBe(0)
    // and a fresh save of the reopened file with no further edits is a no-op
    // (hidden passthrough blocks are skipped, like a session's toSaveBlocks)
    const reopened = await parseDocx(first)
    const resave = await saveDocx(
      reopened,
      reopened.blocks
        .filter((block) => !block.hidden && block.docxIndex !== null)
        .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
    )
    expect(Buffer.compare(Buffer.from(first), Buffer.from(resave))).toBe(0)
  })

  it('a custom heading style with no level edit keeps its pStyle (round-trip)', async () => {
    const doc = await fixture()
    // text-edit shape: structure unchanged, raw pPr passthrough
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[3, retag(doc.blocks[3]!, {})]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[3]!.type).toBe('heading')
    expect(reopened.blocks[3]!.level).toBe(2)
    expect(reopened.blocks[3]!.styleId).toBe('MyHead2')
    expect(reopened.internal.documentXml).toContain('<w:pStyle w:val="MyHead2"/>')
  })

  it('a heading with a direct w:outlineLvl keeps its bytes when unedited', async () => {
    const doc = await fixture()
    expect(doc.blocks[4]!.level).toBe(1) // direct outlineLvl wins over the Heading2 style
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[4, retag(doc.blocks[4]!, {})]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[4]!.type).toBe('heading')
    expect(reopened.blocks[4]!.level).toBe(1)
    expect(reopened.blocks[4]!.styleId).toBe('Heading2')
  })

  it('retagging a direct-outlineLvl heading reopens at the new level', async () => {
    const doc = await fixture()
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[4, retag(doc.blocks[4]!, { level: 3 })]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[4]!.type).toBe('heading')
    expect(reopened.blocks[4]!.level).toBe(3)
  })

  it('an explicit-body-text paragraph (pStyle Heading1 + outlineLvl 9) keeps its style', async () => {
    const doc = await fixture()
    expect(doc.blocks[6]!.type).toBe('paragraph') // outlineLvl 9 = body text, not a heading
    expect(doc.blocks[6]!.styleId).toBe('Heading1')
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[6, retag(doc.blocks[6]!, {})]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[6]!.type).toBe('paragraph')
    expect(reopened.blocks[6]!.styleId).toBe('Heading1')
    expect(reopened.internal.documentXml).toContain('<w:outlineLvl w:val="9"/>')
  })
})

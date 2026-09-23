// BUG-1631: setHeadingLevel to a level whose style the document does not
// define reported success while the save silently kept the old style (the raw
// pPr passthrough kept the stale w:pStyle, and the rebuild path stripped the
// pStyle entirely). The honest contract: when no HeadingN style exists for
// the new level, the retag is carried by a direct w:outlineLvl — Word's own
// outline-level override — so the edit survives save+reopen, the previous
// style (and its formatting) stays, and the mcp op can disclose the
// compromise in its result. These tests pin every retag path (raw pPr
// passthrough, rebuild, outline-only) around a styles.xml that only knows
// Heading1/Heading2, plus the with-style control.
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

// audit scenario corpus: styles.xml knows Heading1/Heading2 only — no Heading3
const NO_H3_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
  '</w:styles>'

const EXTRA_H3 =
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>'

/** h1, style-backed h2, plain tail, outline-only h2 (no pStyle), outline-only h2 (pStyle Normal) */
const BODY = [
  p('<w:pStyle w:val="Heading1"/>', 'H1'),
  p('<w:pStyle w:val="Heading2"/>', 'H2 retag target'),
  '<w:p><w:r><w:t>tail</w:t></w:r></w:p>',
  p('<w:outlineLvl w:val="1"/>', 'outline-only h2'),
  p('<w:pStyle w:val="Normal"/><w:outlineLvl w:val="1"/>', 'outline-only h2 over Normal'),
].join('')

const CTX: GenerateContext = {
  headingStyleIds: new Map([
    [1, 'Heading1'],
    [2, 'Heading2'],
  ]),
  headingLevelOfStyles: new Map<string, number | undefined>([
    ['Heading1', 1],
    ['Heading2', 2],
    ['Normal', undefined],
  ]),
  allocateHyperlinkRel: () => 'rId1',
}

describe('generateParagraphXml: retag to a level with no style (BUG-1631)', () => {
  it('raw passthrough keeps the old pStyle and authors a direct outline level', () => {
    const rawPPr =
      '<w:pPr><w:pStyle w:val="Heading2"/><w:keepNext/><w:spacing w:after="120"/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    // the old style survives (formatting preserved) and the level rides on
    // the direct w:outlineLvl, inserted in CT_PPr position after the other
    // children
    expect(xml).toContain(
      '<w:pPr><w:pStyle w:val="Heading2"/><w:keepNext/><w:spacing w:after="120"/>' +
        '<w:outlineLvl w:val="2"/></w:pPr>',
    )
  })

  it('raw passthrough inserts before a paragraph-mark rPr', () => {
    const rawPPr = '<w:pPr><w:pStyle w:val="Heading2"/><w:rPr><w:i/></w:rPr></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain(
      '<w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="2"/><w:rPr><w:i/></w:rPr></w:pPr>',
    )
  })

  it('raw passthrough keeps a pPrChange revision snapshot as the tail', () => {
    const rawPPr =
      '<w:pPr><w:pStyle w:val="Heading2"/>' +
      '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">' +
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading2"/><w:outlineLvl w:val="2"/>')
    expect(xml).toContain(
      '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>',
    )
  })

  it('rewrites a stale direct outlineLvl instead of duplicating it', () => {
    const rawPPr = '<w:pPr><w:pStyle w:val="Heading2"/><w:outlineLvl w:val="1"/></w:pPr>'
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr, runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:outlineLvl w:val="2"/>')
    expect(xml.match(/<w:outlineLvl/g)).toHaveLength(1)
  })

  it('expands a self-closing empty raw pPr', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', rawPPr: '<w:pPr/>', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pPr><w:outlineLvl w:val="2"/></w:pPr>')
  })

  it('rebuild keeps the stale styleId and authors the level directly', () => {
    const xml = generateParagraphXml(
      { type: 'heading', level: 3, styleId: 'Heading2', runs: [{ text: 'x' }] },
      CTX,
    )
    expect(xml).toContain('<w:pStyle w:val="Heading2"/>')
    expect(xml).toContain('<w:outlineLvl w:val="2"/>')
  })

  it('a fresh heading at a style-less level becomes a genuine outline-only heading', () => {
    const xml = generateParagraphXml({ type: 'heading', level: 3, runs: [{ text: 'x' }] }, CTX)
    expect(xml).not.toContain('w:pStyle')
    expect(xml).toContain('<w:outlineLvl w:val="2"/>')
  })
})

// ---- save/reopen round-trips driven the way an editor session does: the
// retag mutates type/level but neither styleId nor the raw pPr passthrough ----

/** port of the editor-session edit contract (editEntry + toSaveBlocks) */
function retag(block: Block, level: number | 0): GeneratedBlock {
  const gen: GeneratedBlock =
    level === 0
      ? { type: 'paragraph', runs: (block.runs ?? []).map((run) => ({ ...run })) }
      : {
          type: 'heading',
          level,
          ...(block.outlineOnly ? { outlineOnly: true } : {}),
          ...(block.styleId !== undefined ? { styleId: block.styleId } : {}),
          runs: (block.runs ?? []).map((run) => ({ ...run })),
        }
  const structureSame = block.type === gen.type && (block.styleId ?? null) === (gen.styleId ?? null)
  if (structureSame && block.rawPPr !== undefined) gen.rawPPr = block.rawPPr
  return gen
}

async function fixture(extraStyles?: string) {
  return parseDocx(
    await buildDocx({
      stylesXml: extraStyles
        ? NO_H3_STYLES.replace('</w:styles>', `${extraStyles}</w:styles>`)
        : NO_H3_STYLES,
      bodyXml: BODY,
    }),
  )
}

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

describe('heading retag without a style survives save and reopen (BUG-1631)', () => {
  it('style-backed h2 -> h3: reopens at level 3 with the style kept', async () => {
    const doc = await fixture()
    expect(doc.blocks[1]!.level).toBe(2)
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[1, retag(doc.blocks[1]!, 3)]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[1]!.type).toBe('heading')
    expect(reopened.blocks[1]!.level).toBe(3)
    expect(reopened.blocks[1]!.styleId).toBe('Heading2')
    expect(reopened.internal.documentXml).toContain('<w:outlineLvl w:val="2"/>')
  })

  it('style-backed h2 -> h3: saving the retag twice is byte-identical', async () => {
    const doc = await fixture()
    const edits = new Map([[1, retag(doc.blocks[1]!, 3)]])
    const first = await saveDocx(doc, finalBlocksWith(doc, edits))
    const second = await saveDocx(doc, finalBlocksWith(doc, edits))
    expect(Buffer.compare(Buffer.from(first), Buffer.from(second))).toBe(0)
    // and a fresh save of the reopened file with no further edits is a no-op
    const reopened = await parseDocx(first)
    const resave = await saveDocx(
      reopened,
      reopened.blocks
        .filter((block) => !block.hidden && block.docxIndex !== null)
        .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
    )
    expect(Buffer.compare(Buffer.from(first), Buffer.from(resave))).toBe(0)
  })

  it('outline-only h2 -> h3: rewrites the direct outline level (was stale passthrough)', async () => {
    const doc = await fixture()
    expect(doc.blocks[3]!.outlineOnly).toBe(true)
    expect(doc.blocks[3]!.level).toBe(2)
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[3, retag(doc.blocks[3]!, 4)]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[3]!.type).toBe('heading')
    expect(reopened.blocks[3]!.level).toBe(4)
    expect(reopened.blocks[3]!.outlineOnly).toBe(true)
  })

  it('outline-only h2 over a style: retag keeps the style and reopens at the new level', async () => {
    const doc = await fixture()
    expect(doc.blocks[4]!.outlineOnly).toBe(true)
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[4, retag(doc.blocks[4]!, 3)]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[4]!.type).toBe('heading')
    expect(reopened.blocks[4]!.level).toBe(3)
    expect(reopened.blocks[4]!.styleId).toBe('Normal')
  })

  it('demotion to a body paragraph still drops the heading (level 0 unaffected)', async () => {
    const doc = await fixture()
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[1, retag(doc.blocks[1]!, 0)]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[1]!.type).toBe('paragraph')
    expect(reopened.blocks[1]!.level).toBeUndefined()
  })
})

describe('control: a level WITH a style still retargets the pStyle', () => {
  it('h2 -> h3 with Heading3 defined: pStyle rewritten, no override needed', async () => {
    const doc = await fixture(EXTRA_H3)
    expect(doc.blocks[1]!.level).toBe(2)
    const saved = await saveDocx(
      doc,
      finalBlocksWith(doc, new Map([[1, retag(doc.blocks[1]!, 3)]])),
    )
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[1]!.type).toBe('heading')
    expect(reopened.blocks[1]!.level).toBe(3)
    expect(reopened.blocks[1]!.styleId).toBe('Heading3')
    expect(reopened.internal.documentXml).toContain('<w:pStyle w:val="Heading3"/>')
    expect(reopened.internal.documentXml).not.toContain('<w:outlineLvl w:val="2"/></w:pPr>')
  })
})

import { describe, expect, it } from 'vitest'
import {
  buildTocInstruction,
  generateCaptionXml,
  generateTocFieldXml,
  parseDocx,
  parseTocInstruction,
  type TocEntry,
  type TocFieldOptions,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const ENTRIES: TocEntry[] = [
  { level: 1, text: 'Chapter One', pageNo: 1 },
  { level: 2, text: 'Section A', pageNo: 2 },
  { level: 3, text: 'Deep', pageNo: 3 },
]

const instructionOf = (fragments: string[]): string =>
  fragments
    .map((xml) => /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(xml)?.[1] ?? '')
    .join('')
    .trim()

describe('TOC field instruction switches', () => {
  it('default options keep the canonical Word instruction (byte-compat)', () => {
    const fragments = generateTocFieldXml(ENTRIES)
    expect(instructionOf(fragments)).toBe('TOC \\o "1-3" \\h \\z \\u')
  })

  it('levels range narrows \\o and drops deeper cached entries', () => {
    const fragments = generateTocFieldXml(ENTRIES, { levels: 2 })
    expect(instructionOf(fragments)).toBe('TOC \\o "1-2" \\h \\z \\u')
    expect(fragments).toHaveLength(2)
    expect(fragments[0]).toContain('Chapter One')
    expect(fragments.join('')).not.toContain('Deep')
    // deeper entries still get their own TOC level style
    expect(fragments[1]).toContain('<w:pStyle w:val="TOC2"/>')
  })

  it('\\n hides the tab leader and page number runs', () => {
    const fragments = generateTocFieldXml(ENTRIES, { hidePageNumbers: true })
    expect(instructionOf(fragments)).toBe('TOC \\o "1-3" \\n \\h \\z \\u')
    expect(fragments[0]).not.toContain('<w:tab/>')
    expect(fragments[0]).not.toContain('<w:t>1</w:t>')
    expect(fragments[0]).toContain('Chapter One')
  })

  it('hyperlinks off drops the \\h switch', () => {
    expect(instructionOf(generateTocFieldXml(ENTRIES, { hyperlinks: false }))).toBe(
      'TOC \\o "1-3" \\z \\u',
    )
  })

  it('source styles replace \\o/\\u with \\t (Word behavior)', () => {
    expect(instructionOf(generateTocFieldXml(ENTRIES, { styles: 'Chapter 1,1,Appendix,2' }))).toBe(
      'TOC \\h \\z \\t "Chapter 1,1,Appendix,2"',
    )
  })

  it('buildTocInstruction clamps the level range into 1-9', () => {
    expect(buildTocInstruction({ levels: 42 })).toContain('\\o "1-9"')
    expect(buildTocInstruction({ levels: 0 })).toContain('\\o "1-1"')
  })
})

describe('table of figures field', () => {
  const tofEntries: TocEntry[] = [
    { level: 1, text: 'Figure 1 System architecture', pageNo: 2, anchor: '_Ref100000001' },
    { level: 1, text: 'Figure 2 Data flow', pageNo: 5, anchor: '_Ref100000002' },
  ]

  it('emits the Word table-of-figures instruction with the SEQ label', () => {
    const fragments = generateTocFieldXml(tofEntries, { seqIdentifier: 'Figure' })
    expect(instructionOf(fragments)).toBe('TOC \\h \\z \\c "Figure"')
    expect(fragments).toHaveLength(2)
  })

  it('wraps each entry in a w:hyperlink to the caption anchor', () => {
    const fragments = generateTocFieldXml(tofEntries, { seqIdentifier: 'Figure' })
    expect(fragments[0]).toContain('<w:hyperlink w:anchor="_Ref100000001" w:history="1">')
    expect(fragments[0]).toContain('Figure 1 System architecture')
    expect(fragments[0]).toContain('<w:tab/>')
    expect(fragments[0]).toContain('<w:t>2</w:t>')
    // field begin/separate stay outside the hyperlink, the end closes after it
    expect(fragments[0].indexOf('w:fldCharType="separate"')).toBeLessThan(
      fragments[0].indexOf('<w:hyperlink'),
    )
    expect(fragments[1].endsWith('<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')).toBe(true)
  })

  it('round-trips through parseDocx as level-1 tocLines with anchors', async () => {
    const caption1 = generateCaptionXml('Figure', 1, 'System architecture', '_Ref100000001')
    const caption2 = generateCaptionXml('Figure', 2, 'Data flow', '_Ref100000002')
    const tof = generateTocFieldXml(tofEntries, { seqIdentifier: 'Figure' })
    const doc = await parseDocx(await buildDocx({ bodyXml: tof.join('') + caption1 + caption2 }))
    const lines = doc.blocks.filter((b) => b.fieldDisplay?.kind === 'tocLine')
    expect(lines).toHaveLength(2)
    expect(lines[0].fieldDisplay).toMatchObject({
      kind: 'tocLine',
      left: 'Figure 1 System architecture',
      right: '2',
      level: 1,
      anchor: '_Ref100000001',
    })
    expect(lines[1].fieldDisplay).toMatchObject({
      left: 'Figure 2 Data flow',
      right: '5',
      anchor: '_Ref100000002',
    })
  })

  it('round-trips \\n entries as title-only (noPage) lines', async () => {
    const fragments = generateTocFieldXml(ENTRIES, { hidePageNumbers: true })
    const doc = await parseDocx(await buildDocx({ bodyXml: fragments.join('') }))
    const lines = doc.blocks.filter((b) => b.fieldDisplay?.kind === 'tocLine')
    expect(lines).toHaveLength(3)
    expect(lines.every((b) => b.fieldDisplay?.noPage === true)).toBe(true)
    expect(lines[0].fieldDisplay?.left).toBe('Chapter One')
    expect(lines[0].fieldDisplay?.right).toBe('')
  })
})

describe('parseTocInstruction', () => {
  const cases: Array<[string, TocFieldOptions]> = [
    ['TOC \\o "1-3" \\h \\z \\u', { levels: 3, hyperlinks: true }],
    ['TOC \\o "1-2" \\n \\h \\z \\u', { levels: 2, hidePageNumbers: true, hyperlinks: true }],
    ['TOC \\h \\z \\t "Chapter 1,1"', { styles: 'Chapter 1,1', hyperlinks: true }],
    ['TOC \\o "1-1" \\z \\u', { levels: 1, hyperlinks: false }],
    ['TOC \\h \\z \\c "Figure"', { seqIdentifier: 'Figure', hyperlinks: true }],
  ]
  for (const [instr, expected] of cases) {
    it(`parses ${instr}`, () => {
      expect(parseTocInstruction(instr)).toEqual(expected)
    })
  }

  it('regenerating from a parsed instruction keeps the switches (update loop)', () => {
    for (const [instr] of cases) {
      const options = parseTocInstruction(instr)
      expect(buildTocInstruction(options)).toBe(instr)
    }
  })
})

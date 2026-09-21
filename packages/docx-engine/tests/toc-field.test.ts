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
    [
      'TOC \\o "1-3" \\n 2-4 \\h \\z \\u',
      { levels: 3, hidePageNumbersFrom: 2, hidePageNumbersTo: 4, hyperlinks: true },
    ],
    ['TOC \\h \\z \\t "Chapter 1,1"', { styles: 'Chapter 1,1', hyperlinks: true }],
    // \o + \t union: the update path collects headings AND style-mapped
    // paragraphs, and the instruction keeps both switches (BUG-1012)
    [
      'TOC \\o "1-3" \\h \\z \\t "Chapter 1,1"',
      { levels: 3, styles: 'Chapter 1,1', hyperlinks: true },
    ],
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

describe('parseTocInstruction alternative spellings (BUG-1010)', () => {
  // Word writes ` TOC \c "Figure" ` (space + double quotes); other producers
  // emit \c without a space, single-quoted, or bare — the reader takes them
  // all so an update keeps the field a table of figures instead of silently
  // regenerating it as a heading TOC.
  it('reads \\c without a space, single-quoted, and bare', () => {
    expect(parseTocInstruction('TOC \\c"Figure" \\h')).toEqual({
      seqIdentifier: 'Figure',
      hyperlinks: true,
    })
    expect(parseTocInstruction("TOC \\c 'Figure' \\h")).toEqual({
      seqIdentifier: 'Figure',
      hyperlinks: true,
    })
    expect(parseTocInstruction('TOC \\c Figure \\h')).toEqual({
      seqIdentifier: 'Figure',
      hyperlinks: true,
    })
  })

  it('keeps embedded spaces inside quotes', () => {
    expect(parseTocInstruction('TOC \\c "My Figure" \\h')).toEqual({
      seqIdentifier: 'My Figure',
      hyperlinks: true,
    })
  })

  it('reads \\t and \\o without quotes or padding spaces', () => {
    expect(parseTocInstruction('TOC \\t"Chapter 1,1" \\h')).toEqual({
      styles: 'Chapter 1,1',
      hyperlinks: true,
    })
    expect(parseTocInstruction('TOC \\o 1-2')).toEqual({ levels: 2, hyperlinks: false })
    expect(parseTocInstruction("TOC \\o'1-3' \\h")).toEqual({ levels: 3, hyperlinks: true })
  })
})

describe('parseTocInstruction does not swallow the next switch (BUG-1109)', () => {
  it('reads \\c followed by another switch as an empty identifier', () => {
    // \S+ used to match the backslash too: `TOC \c \h` parsed "\h" as the SEQ
    // label and regeneration wrote a garbage `TOC \h \z \c "\h"`
    expect(parseTocInstruction('TOC \\c \\h \\z')).toEqual({ hyperlinks: true })
    expect(parseTocInstruction('TOC \\c\\h')).toEqual({ hyperlinks: true })
    expect(parseTocInstruction('TOC \\h \\z \\c')).toEqual({ hyperlinks: true })
    expect(parseTocInstruction('TOC \\c \\h \\z')).not.toHaveProperty('seqIdentifier')
  })

  it('bare words still work but stop at a backslash (BUG-1010 spellings intact)', () => {
    expect(parseTocInstruction('TOC \\c Figure \\h')).toEqual({
      seqIdentifier: 'Figure',
      hyperlinks: true,
    })
    expect(parseTocInstruction('TOC \\cFigure \\h')).toEqual({
      seqIdentifier: 'Figure',
      hyperlinks: true,
    })
  })

  it('a no-space ranged \\n parses per level, and \\h survives a packed switch', () => {
    // `TOC \n2-4` used to widen to a full \n — the regenerated instruction
    // hid every level's page numbers instead of levels 2-4 only (BUG-1012)
    expect(parseTocInstruction('TOC \\o "1-3" \\n2-4 \\h')).toEqual({
      levels: 3,
      hidePageNumbersFrom: 2,
      hidePageNumbersTo: 4,
      hyperlinks: true,
    })
    expect(parseTocInstruction('TOC \\o "1-3" \\n 2-4 \\h \\z \\u')).toEqual({
      levels: 3,
      hidePageNumbersFrom: 2,
      hidePageNumbersTo: 4,
      hyperlinks: true,
    })
    expect(parseTocInstruction('TOC \\n')).toEqual({ hidePageNumbers: true, hyperlinks: false })
    expect(parseTocInstruction('TOC \\h\\z')).toEqual({ hyperlinks: true })
  })
})

describe('ranged \\n keeps page numbers per level (BUG-1012)', () => {
  it('regenerating from a parsed ranged \\n round-trips the range', () => {
    expect(buildTocInstruction({ levels: 3, hidePageNumbersFrom: 2, hidePageNumbersTo: 4 })).toBe(
      'TOC \\o "1-3" \\n 2-4 \\h \\z \\u',
    )
    const parsed = parseTocInstruction('TOC \\o "1-3" \\n 2-4 \\h \\z \\u')
    expect(buildTocInstruction(parsed)).toBe('TOC \\o "1-3" \\n 2-4 \\h \\z \\u')
  })

  it('levels inside the range lose the tab and page number; levels outside keep them', () => {
    const fragments = generateTocFieldXml(ENTRIES, {
      levels: 3,
      hidePageNumbersFrom: 2,
      hidePageNumbersTo: 4,
    })
    expect(instructionOf(fragments)).toBe('TOC \\o "1-3" \\n 2-4 \\h \\z \\u')
    // level 1 (Chapter One): page number kept
    expect(fragments[0]).toContain('Chapter One')
    expect(fragments[0]).toContain('<w:tab/>')
    expect(fragments[0]).toContain('<w:t>1</w:t>')
    // levels 2-3 (Section A, Deep): title-only cached entries
    expect(fragments[1]).not.toContain('<w:tab/>')
    expect(fragments[1]).not.toContain('<w:t>2</w:t>')
    expect(fragments[2]).not.toContain('<w:tab/>')
  })

  it('round-trips through parseDocx: in-range lines are noPage, level 1 keeps its page', async () => {
    const fragments = generateTocFieldXml(ENTRIES, {
      levels: 3,
      hidePageNumbersFrom: 2,
      hidePageNumbersTo: 3,
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: fragments.join('') }))
    const lines = doc.blocks.filter((b) => b.fieldDisplay?.kind === 'tocLine')
    expect(lines).toHaveLength(3)
    expect(lines[0].fieldDisplay).toMatchObject({ left: 'Chapter One', right: '1', level: 1 })
    expect(lines[0].fieldDisplay?.noPage).toBeUndefined()
    expect(lines[1].fieldDisplay).toMatchObject({ left: 'Section A', right: '', noPage: true })
    expect(lines[2].fieldDisplay?.noPage).toBe(true)
  })
})

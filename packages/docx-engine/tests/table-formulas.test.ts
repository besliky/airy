import { describe, expect, it } from 'vitest'
import {
  evaluateFormulaInGrid,
  FORMULA_DIV_ZERO_ERROR,
  FORMULA_EMPTY_ERROR,
  FORMULA_SYNTAX_ERROR,
  formatNumericPicture,
  parseCellNumber,
  parseFormulaInstruction,
  proposeTableFormula,
} from '@airy-office/docx-engine'
import { generateTableModelXml, parseDocx } from '@airy-office/docx-engine'
import { buildDocx } from './helpers/build-docx'

/** 3x3 grid for direction/ref tests; one text cell cuts the B column scan */
const GRID = [
  ['1', '2', '3'],
  ['4', 'x', '6'],
  ['7', '8', '9'],
]

const gridOf = (texts: string[][]) => ({ texts })

describe('table formula direction operands', () => {
  it('sums ABOVE/BELOW/LEFT/RIGHT from the neighbor cell while cells hold numbers', () => {
    const texts = [
      ['1', '2', '3'],
      ['4', '5', '6'],
      ['7', '8', '9'],
    ]
    const g = gridOf(texts)
    // ABOVE at (2,1) = 5+2; BELOW at (0,1) = 5+8; LEFT at (1,2) = 5+4; RIGHT at (1,0) = 5+6
    expect(evaluateFormulaInGrid('=SUM(ABOVE)', g, 2, 1)).toBe('7')
    expect(evaluateFormulaInGrid('=SUM(BELOW)', g, 0, 1)).toBe('13')
    expect(evaluateFormulaInGrid('=SUM(LEFT)', g, 1, 2)).toBe('9')
    expect(evaluateFormulaInGrid('=SUM(RIGHT)', g, 1, 0)).toBe('11')
  })

  it('stops the scan at the first blank or text cell like Word', () => {
    const g = gridOf([
      ['100', 'x', '1'],
      ['200', '', '2'],
      ['a', '3', '3'],
    ])
    // blank at (1,1) cuts off everything above it
    expect(evaluateFormulaInGrid('=SUM(ABOVE)', g, 2, 1)).toBe(FORMULA_EMPTY_ERROR)
    // text cell stops the LEFT scan
    expect(evaluateFormulaInGrid('=SUM(LEFT)', g, 2, 2)).toBe('3')
    // column C has clean numbers above: 2 + 1
    expect(evaluateFormulaInGrid('=SUM(ABOVE)', g, 2, 2)).toBe('3')
  })

  it('repeats the empty-column edge as Word’s Undefined Bookmark error', () => {
    const g = gridOf([
      ['', '1'],
      ['', '2'],
    ])
    expect(evaluateFormulaInGrid('=SUM(ABOVE)', g, 2, 0)).toBe(FORMULA_EMPTY_ERROR)
    expect(evaluateFormulaInGrid('=AVERAGE(ABOVE)', g, 0, 0)).toBe(FORMULA_EMPTY_ERROR)
  })

  it('ignores repeating-header rows during direction scans (Word skips heading rows)', () => {
    const g = {
      texts: [
        ['5', '12'],
        ['1', '2'],
        ['', ''],
      ],
      isHeaderRow: (r: number) => r === 0,
    }
    // without the header rule the sum would pick up the header 5 → 6
    expect(evaluateFormulaInGrid('=SUM(ABOVE)', g, 2, 0)).toBe('1')
    expect(evaluateFormulaInGrid('=SUM(ABOVE)', g, 2, 1)).toBe('2')
  })

  it('supports cell references, ranges and the common function set', () => {
    const g = gridOf(GRID)
    expect(evaluateFormulaInGrid('=SUM(A1:B2)', g, 2, 2)).toBe('7') // 1+2+4, text skipped
    expect(evaluateFormulaInGrid('=A1*2', g, 2, 2)).toBe('2')
    expect(evaluateFormulaInGrid('=AVERAGE(B1:B3)', g, 2, 2)).toBe('5') // 2 and 8
    expect(evaluateFormulaInGrid('=MIN(A1:A3)', g, 2, 2)).toBe('1')
    expect(evaluateFormulaInGrid('=MAX(C1:C2)', g, 2, 2)).toBe('6')
    expect(evaluateFormulaInGrid('=COUNT(A1:B3)', g, 2, 2)).toBe('5')
    expect(evaluateFormulaInGrid('=PRODUCT(A1:C1)', g, 2, 2)).toBe('6')
    expect(evaluateFormulaInGrid('=ROUND(2.345,2)', g, 2, 2)).toBe('2.35')
    expect(evaluateFormulaInGrid('=ROUND(-2.5)', g, 2, 2)).toBe('-3')
    expect(evaluateFormulaInGrid('=MOD(5,3)', g, 2, 2)).toBe('2')
    expect(evaluateFormulaInGrid('=ABS(-4)+INT(1.8)+SIGN(-9)', g, 2, 2)).toBe('4')
    expect(evaluateFormulaInGrid('=IF(A1=1,10%,20%)', g, 2, 2)).toBe('0.1')
    expect(evaluateFormulaInGrid('=DEFINED(1/1)+DEFINED(1/0)', g, 2, 2)).toBe('1')
    expect(evaluateFormulaInGrid('=1+', g, 2, 2)).toBe(FORMULA_SYNTAX_ERROR)
    expect(evaluateFormulaInGrid('=1/0', g, 2, 2)).toBe(FORMULA_DIV_ZERO_ERROR)
  })
})

describe('formula instruction parsing and numeric pictures', () => {
  it('splits the expression from the \\# picture switch (Word spells \\# "picture")', () => {
    expect(parseFormulaInstruction('=SUM(ABOVE)')).toEqual({ expr: 'SUM(ABOVE)', picture: null })
    expect(parseFormulaInstruction(' =SUM(ABOVE) \\# "#,##0.00" ')).toEqual({
      expr: 'SUM(ABOVE)',
      picture: '#,##0.00',
    })
    expect(parseFormulaInstruction('=1+2 \\* MERGEFORMAT')).toEqual({
      expr: '1+2',
      picture: null,
    })
    expect(parseFormulaInstruction('SUM(ABOVE)')).toBeNull()
  })

  it('parses cell numbers strictly (blank and text cells are not numeric)', () => {
    expect(parseCellNumber(' 12.5 ')).toBe(12.5)
    expect(parseCellNumber('-3')).toBe(-3)
    expect(parseCellNumber('1 000')).toBe(1000)
    expect(parseCellNumber('1,000')).toBeNull()
    expect(parseCellNumber('N/A')).toBeNull()
    expect(parseCellNumber('')).toBeNull()
  })

  it('formats the minimal Word numeric picture set', () => {
    expect(formatNumericPicture(1234.5, '#,##0.00')).toBe('1,234.50')
    expect(formatNumericPicture(0.5, '0')).toBe('1')
    expect(formatNumericPicture(0.25, '0.00')).toBe('0.25')
    expect(formatNumericPicture(1234, '#,##0')).toBe('1,234')
    expect(formatNumericPicture(0.1234, '0%')).toBe('12%')
    expect(formatNumericPicture(1234.5, '$#,##0.00')).toBe('$1,234.50')
    // positive;negative sections spell the sign themselves
    expect(formatNumericPicture(-1234.5, '#,##0.00;(#,##0.00)')).toBe('(1,234.50)')
    expect(formatNumericPicture(-1234.5, '#,##0.00')).toBe('-1,234.50')
    expect(formatNumericPicture(0, '0;-0;"—"')).toBe('—')
    expect(formatNumericPicture(2.5, '0.##')).toBe('2.5')
  })

  it('applies the picture switch to the evaluated result', () => {
    const g = gridOf([
      ['10.5', '0.25'],
      ['20.25', '0.5'],
    ])
    expect(evaluateFormulaInGrid('=SUM(ABOVE) \\# "#,##0.00"', g, 2, 0)).toBe('30.75')
    expect(evaluateFormulaInGrid('=SUM(ABOVE) \\# 0%', g, 2, 1)).toBe('75%')
  })
})

/** Word's own spelling: a w:fldSimple as a direct w:p child with a cached result run */
const formulaCell = (instr: string, cached: string) =>
  `<w:tc><w:p><w:fldSimple w:instr="${instr}">` +
  `<w:r><w:t>${cached}</w:t></w:r></w:fldSimple></w:p></w:tc>`

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

describe('fldSimple formula parse/serialize inside table cells', () => {
  it('parses a formula cell into a formulaField run carrying the cached result', async () => {
    const bodyXml =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell('10')}${cell('20')}</w:tr>` +
      `<w:tr>${formulaCell(' =SUM(ABOVE) \\# &quot;#,##0.00&quot; ', '30.00')}${cell('7')}</w:tr>` +
      '</w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const table = parsed.blocks.find((b) => b.type === 'table')
    expect(table && table.type === 'table').toBe(true)
    if (!table || table.type !== 'table' || !table.table) return
    const run = table.table.rows[1][0].richParas?.[0].runs[0]
    expect(run?.formulaField).toBe('=SUM(ABOVE) \\# "#,##0.00"')
    expect(run?.text).toBe('30.00')
  })

  it('re-emits the w:fldSimple with its instruction and cached result on save', async () => {
    const bodyXml =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell('1')}${cell('2')}</w:tr>` +
      `<w:tr>${formulaCell('=SUM(ABOVE)', '3')}${cell('x')}</w:tr>` +
      '</w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const table = parsed.blocks.find((b) => b.type === 'table')
    if (!table || table.type !== 'table' || !table.table) return
    const xml = generateTableModelXml(table.table)
    expect(xml).toContain('<w:fldSimple w:instr="=SUM(ABOVE)"')
    expect(xml).toContain('>3</w:t>')
    // plain cells stay plain
    expect(xml).toContain('>x</w:t>')
  })

  it('round-trips a complex-field formula into the fldSimple form and back', async () => {
    const bodyXml =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> =SUM(ABOVE) </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>42</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const table = parsed.blocks.find((b) => b.type === 'table')
    if (!table || table.type !== 'table' || !table.table) return
    const run = table.table.rows[0][0].richParas?.[0].runs[0]
    expect(run?.formulaField).toBe('=SUM(ABOVE)')
    expect(run?.text).toBe('42')
  })
})

/** BUG-1756 shape: the complex/fldSimple field carries NO cached result; the
 *  displayed value follows the field as a plain run (common generator output) */
const EMPTY_COMPLEX_FORMULA_CELL =
  '<w:tc><w:p>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> =SUM(ABOVE) </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t>600</w:t></w:r>' +
  '</w:p></w:tc>'

const wrapTable = (cellXml: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cellXml}</w:tr></w:tbl>`

const tableRun = async (bodyXml: string) => {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const table = parsed.blocks.find((b) => b.type === 'table')
  if (!table || table.type !== 'table' || !table.table) return null
  return table.table.rows[0][0].richParas?.[0].runs ?? []
}

describe('empty-cache formula fields fold the trailing value run (BUG-1756)', () => {
  it('folds the plain numeric run after an empty complex field into the formula run', async () => {
    const runs = await tableRun(wrapTable(EMPTY_COMPLEX_FORMULA_CELL))
    expect(runs).toHaveLength(1)
    expect(runs?.[0]?.text).toBe('600')
    expect(runs?.[0]?.formulaField).toBe('=SUM(ABOVE)')
  })

  it('re-emits one fldSimple with the folded cache and no stray run on save', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: wrapTable(EMPTY_COMPLEX_FORMULA_CELL) }),
    )
    const table = parsed.blocks.find((b) => b.type === 'table')
    if (!table || table.type !== 'table' || !table.table) return
    const xml = generateTableModelXml(table.table)
    expect(xml.match(/<w:fldSimple /g)).toHaveLength(1)
    expect(xml).toContain(
      '<w:fldSimple w:instr="=SUM(ABOVE)"><w:r><w:t xml:space="preserve">600</w:t></w:r></w:fldSimple>',
    )
    expect(xml.match(/>600<\/w:t>/g)).toHaveLength(1)
  })

  it('keeps the single-space placeholder when nothing foldable follows the field', async () => {
    const cellXml =
      '<w:tc><w:p>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> =SUM(ABOVE) </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:t>Total</w:t></w:r>' +
      '</w:p></w:tc>'
    const runs = await tableRun(wrapTable(cellXml))
    expect(runs).toHaveLength(2)
    expect(runs?.[0]?.text).toBe(' ')
    expect(runs?.[0]?.formulaField).toBe('=SUM(ABOVE)')
    expect(runs?.[1]?.text).toBe('Total')
    expect(runs?.[1]?.formulaField).toBeUndefined()
  })

  it('folds under a comment range and keeps the comment anchor on the folded run', async () => {
    const cellXml =
      '<w:tc><w:p><w:commentRangeStart w:id="0"/>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> =SUM(ABOVE) </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:t>600</w:t></w:r>' +
      '<w:commentRangeEnd w:id="0"/>' +
      '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>' +
      '</w:p></w:tc>'
    const runs = await tableRun(wrapTable(cellXml))
    expect(runs).toHaveLength(1)
    expect(runs?.[0]?.text).toBe('600')
    expect(runs?.[0]?.formulaField).toBe('=SUM(ABOVE)')
    expect(runs?.[0]?.commentIds).toEqual(['0'])
    const parsed = await parseDocx(await buildDocx({ bodyXml: wrapTable(cellXml) }))
    const table = parsed.blocks.find((b) => b.type === 'table')
    if (!table || table.type !== 'table' || !table.table) return
    const xml = generateTableModelXml(table.table)
    // the saved cell keeps the comment range around exactly one fldSimple cache
    expect(xml.indexOf('<w:commentRangeStart w:id="0"/>')).toBeLessThan(
      xml.indexOf('<w:fldSimple w:instr="=SUM(ABOVE)"'),
    )
    expect(xml.indexOf('<w:fldSimple w:instr="=SUM(ABOVE)"')).toBeLessThan(
      xml.indexOf('<w:commentRangeEnd w:id="0"/>'),
    )
    expect(xml.match(/>600<\/w:t>/g)).toHaveLength(1)
  })

  it('folds an empty w:fldSimple formula the same way', async () => {
    const cellXml =
      '<w:tc><w:p><w:fldSimple w:instr="=SUM(ABOVE)"></w:fldSimple>' +
      '<w:r><w:t>600</w:t></w:r></w:p></w:tc>'
    const runs = await tableRun(wrapTable(cellXml))
    expect(runs).toHaveLength(1)
    expect(runs?.[0]?.text).toBe('600')
    expect(runs?.[0]?.formulaField).toBe('=SUM(ABOVE)')
  })
})

describe('Word dialog prefill', () => {
  it('proposes SUM(ABOVE) over a numeric column, else SUM(LEFT)', () => {
    const g = gridOf([
      ['', '1', ''],
      ['', '2', ''],
      ['3', '4', ''],
    ])
    expect(proposeTableFormula(g.texts, 2, 1)).toBe('=SUM(ABOVE)')
    expect(proposeTableFormula(g.texts, 2, 2)).toBe('=SUM(LEFT)')
    // Word's default stays =SUM(ABOVE) when neither neighbor is numeric
    expect(proposeTableFormula(g.texts, 2, 0)).toBe('=SUM(ABOVE)')
  })
})

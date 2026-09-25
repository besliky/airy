/**
 * Recalculated values lived only on screen, so a saved file had new
 * inputs and stale outputs — readers without a formula engine (openpyxl data_only,
 * pandas, preview services) silently got wrong numbers. The save now refreshes each
 * formula cell's cached <v> while leaving its <f> untouched.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import { buildStructureFixture } from './fixture-builder'

const SHEET = 'Data'

/** The fixture's D2 is `=SUM(A1:A2)` with a cached value. */
async function saveWithFormulaValues(
  cells: { row: number; column: number; value: string | number | boolean | null }[],
) {
  const mutation = await applyCellEditsToXlsx(
    await buildStructureFixture(),
    // one unrelated value edit so the save has something to do
    [{ sheetName: SHEET, row: 0, column: 0, writeValue: true, cell: { value: 1 } }],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [{ sheetName: SHEET, cells }],
  )
  const zip = await JSZip.loadAsync(mutation.buffer)
  return (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
}

describe('formula cached values', () => {
  it('updates <v> and keeps <f> for a numeric result', async () => {
    const xml = await saveWithFormulaValues([{ row: 1, column: 3, value: 811.4 }])
    const cell = /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('<f>')
    expect(cell).toContain('<v>811.4</v>')
    // exactly one cached value, and no stale one left behind
    expect((cell.match(/<v>/g) ?? []).length).toBe(1)
  })

  it('a text result gets t="str"', async () => {
    const xml = await saveWithFormulaValues([{ row: 1, column: 3, value: 'N/A & more' }])
    const cell = /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('t="str"')
    expect(cell).toContain('<v>N/A &amp; more</v>')
    expect(cell).toContain('<f>')
  })

  it('a null result clears the cached value but keeps the formula', async () => {
    const xml = await saveWithFormulaValues([{ row: 1, column: 3, value: null }])
    const cell = /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''
    expect(cell).toContain('<f>')
    expect(cell).not.toContain('<v>')
  })

  it('non-formula cells and missing cells are left alone', async () => {
    const before = await saveWithFormulaValues([])
    // A1 is a plain number, ZZ99 does not exist
    const after = await saveWithFormulaValues([
      { row: 0, column: 0, value: 999 },
      { row: 98, column: 701, value: 999 },
    ])
    expect(after).toBe(before)
    expect(after).not.toContain('<v>999</v>')
  })
})

// SC-09/OBS-1719: the empty `<v></v>` cache form some generators emit carries
// no information (cache readers see None, fullCalcOnLoad recomputes) — a save
// must not re-serialize it. A real overlay value fills <v>; an unknown cache
// is dropped; a t="str" empty cache is a legitimate cached "" and stays.
describe('formula cells save without the empty-cache form', () => {
  /** The fixture with D2's cache rewritten to the given cell XML, then saved */
  async function saveWithPatchedD2(
    d2Cell: string,
    formulaValues: { row: number; column: number; value: string | number | boolean | null }[] = [],
  ) {
    const zip = await JSZip.loadAsync(await buildStructureFixture())
    const sheet = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet.replace('<c r="D2"><f>SUM(A1:A2)</f><v>3</v></c>', d2Cell),
    )
    const patched = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    const mutation = await applyCellEditsToXlsx(
      new Uint8Array(patched),
      // one unrelated value edit so the save has something to do
      [{ sheetName: SHEET, row: 9, column: 0, writeValue: true, cell: { value: 11 } }],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      formulaValues.length > 0 ? [{ sheetName: SHEET, cells: formulaValues }] : [],
    )
    const out = await JSZip.loadAsync(mutation.buffer)
    return (await out.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
  }

  const d2CellOf = (xml: string) => /<c[^>]*r="D2"[^>]*>[\s\S]*?<\/c>/.exec(xml)?.[0] ?? ''

  it('drops an empty <v></v> when no recalculated value covers the cell', async () => {
    const xml = await saveWithPatchedD2('<c r="D2"><f>SUM(A1:A2)</f><v></v></c>')
    const cell = d2CellOf(xml)
    expect(cell).toContain('<f>SUM(A1:A2)</f>')
    expect(cell).not.toContain('<v>')
  })

  it('drops a self-closing <v/> the same way', async () => {
    const xml = await saveWithPatchedD2('<c r="D2"><f>SUM(A1:A2)</f><v/></c>')
    expect(d2CellOf(xml)).not.toContain('<v')
  })

  it('still writes a recalculated value over the empty cache', async () => {
    const xml = await saveWithPatchedD2('<c r="D2"><f>SUM(A1:A2)</f><v></v></c>', [
      { row: 1, column: 3, value: 811.4 },
    ])
    const cell = d2CellOf(xml)
    expect(cell).toContain('<v>811.4</v>')
    expect((cell.match(/<v>/g) ?? []).length).toBe(1)
  })

  it('keeps a t="str" empty cache (a legitimate cached empty string)', async () => {
    const xml = await saveWithPatchedD2('<c r="D2" t="str"><f>LEFT(A2,0)</f><v></v></c>', [])
    expect(d2CellOf(xml)).toContain('<v></v>')
  })

  it('never touches the cache of a non-formula cell', async () => {
    const xml = await saveWithPatchedD2('<c r="D2"><v></v></c>')
    // D2 has no <f>: it is not a formula cell, the empty value stays verbatim
    expect(d2CellOf(xml)).toBe('<c r="D2"><v></v></c>')
  })
})

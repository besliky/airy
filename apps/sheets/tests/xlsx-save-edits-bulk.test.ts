/**
 * Bulk save-scale guard, split from xlsx-save-edits.test.ts (PERF-1201: the
 * file was the suite's top vitest outlier at 5.4s, ~4s of it here). Bulk
 * actions (paste, sort, move-range) journal every affected cell, so a single
 * save can carry hundreds of thousands of edits. The save path must stay a
 * single pass over the worksheet — a per-edit whole-XML rewrite would take
 * hours at this scale.
 *
 * Scale note: the insert leg keeps the full 100k-cell batch; the overwrite
 * leg runs 25k cells — the dominating real cost (text overwrites on a
 * populated sheet measure ~20ms/1k cells, ~13x the insert path), and a
 * quadratic regression blows the 30s budget loudly at either scale.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, type CellEdit } from '../src/gateway/xlsx-gateway'
import { blankXlsxBuffer } from '../src/gateway/csv-import'

async function entryText(buffer: Buffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file(path)
  if (!entry) throw new Error(`Missing ${path}`)
  return entry.async('text')
}

describe('large edit batches', () => {
  it('applies 100k inserts and 25k text overwrites in one save each', async () => {
    const rowCount = 2_000
    const columnCount = 50
    const inserts: CellEdit[] = []
    for (let row = 0; row < rowCount; row += 1) {
      for (let column = 0; column < columnCount; column += 1) {
        inserts.push({
          sheetName: 'Sheet1',
          row,
          column,
          writeValue: true,
          cell: { value: row * columnCount + column },
        })
      }
    }
    const started = performance.now()
    const insertMutation = await applyCellEditsToXlsx(await blankXlsxBuffer(), inserts)
    const worksheet = await entryText(insertMutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<dimension ref="A1:AX2000"/>')
    expect(worksheet).toContain('<c r="A1"><v>0</v></c>')
    expect(worksheet).toContain(`<c r="AX2000"><v>${rowCount * columnCount - 1}</v></c>`)
    expect([...worksheet.matchAll(/<row /g)]).toHaveLength(rowCount)

    const overwriteRowCount = 500
    const overwrites: CellEdit[] = inserts
      .filter((edit) => edit.row < overwriteRowCount)
      .map((edit) => ({ ...edit, cell: { value: `text ${edit.row}:${edit.column}` } }))
    const overwriteMutation = await applyCellEditsToXlsx(insertMutation.buffer, overwrites)
    const rewritten = await entryText(overwriteMutation.buffer, 'xl/worksheets/sheet1.xml')
    expect(rewritten).toContain('<t xml:space="preserve">text 0:0</t>')
    expect(rewritten).toContain(
      `<t xml:space="preserve">text ${overwriteRowCount - 1}:${columnCount - 1}</t>`,
    )
    expect(rewritten).not.toContain('<v>0</v>')

    // Generous CI budget; the quadratic path this guards against took hours.
    expect(performance.now() - started).toBeLessThan(30_000)
  }, 60_000)
})

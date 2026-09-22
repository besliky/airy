/**
 * BUG-1504 phantom trailing page on bordered sheets: printedHeightPt counted
 * only the text line (fs x 1.25 + 2pt padding), but border-collapse grows
 * every rendered row by its border (a thin edge is ~0.75pt, the gridline net
 * 0.5pt). The accumulated drift (~0.5-0.75pt/row) spilled the table's tail
 * onto an extra, otherwise empty PDF page. The declared <tr> height now
 * includes the row's widest vertical border, keeping the pagination model
 * and the emitted tables in step.
 */
import { describe, expect, it } from 'vitest'

import { buildSheetsPrintPayload, type PrintWorksheet } from '../src/renderer/print-html'
import type { EffectivePageSetup } from '../src/renderer/print-settings'

function payloadSetup(overrides: Partial<EffectivePageSetup> = {}): EffectivePageSetup {
  return {
    orientation: 'portrait',
    paperSize: 9,
    scale: 100,
    fitToWidth: 0,
    fitToHeight: 0,
    fitToPage: false,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    printGridlines: false,
    printHeadings: false,
    printAreas: [],
    printTitles: null,
    header: null,
    footer: null,
    firstPage: null,
    evenPages: null,
    headerFooterScaleWithDoc: true,
    headerFooterPictures: [],
    ...overrides,
  }
}

/// A one-column worksheet (text cells heighten every row to the text line:
/// 11pt x 1.25 + 2pt = 15.75pt) with a per-cell bottom style.
function borderedWorksheet(style: Record<string, unknown> | null): PrintWorksheet {
  return {
    getSheetName: () => 'Bordered',
    getSheet: () => ({
      getRowVisible: () => true,
      getColVisible: () => true,
    }),
    getLastRow: () => 0,
    getLastColumn: () => 0,
    getRowHeight: () => 20, // 20px = 15pt saved
    getColumnWidth: () => 100,
    getMergedRanges: () => [],
    getRange: (() => ({
      getDisplayValues: () => [['x']],
      getValues: () => [[42]],
      getCellStyleData: () => style,
    })) as PrintWorksheet['getRange'],
  }
}

describe('printedHeightPt collapsed-border contribution', () => {
  it('declares the thin border on top of the text line height', () => {
    // 15.75pt text line + 0.75pt thin bottom border.
    const payload = buildSheetsPrintPayload(
      [
        {
          worksheet: borderedWorksheet({ bd: { b: { s: 1 } } }),
          printAreas: [],
          printTitles: null,
        },
      ],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    expect(payload.html).toContain('<tr style="height:16.5pt">')
  })

  it('counts the widest vertical edge (medium beats thin)', () => {
    // 15.75pt text line + 1.5pt medium border.
    const payload = buildSheetsPrintPayload(
      [
        {
          worksheet: borderedWorksheet({ bd: { t: { s: 1 }, b: { s: 8 } } }),
          printAreas: [],
          printTitles: null,
        },
      ],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    expect(payload.html).toContain('<tr style="height:17.25pt">')
  })

  it('counts the gridline net when gridlines print without cell borders', () => {
    // 15.75pt text line + 0.5pt gridline border.
    const payload = buildSheetsPrintPayload(
      [{ worksheet: borderedWorksheet(null), printAreas: [], printTitles: null }],
      payloadSetup({ printGridlines: true }),
      'Book.pdf',
      'S',
    )
    expect(payload.html).toContain('<tr style="height:16.25pt">')
  })

  it('keeps the plain text-line height without borders or gridlines', () => {
    const payload = buildSheetsPrintPayload(
      [{ worksheet: borderedWorksheet(null), printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    expect(payload.html).toContain('<tr style="height:15.75pt">')
  })

  it('a cell borders the whole row even when a neighbour is unstyled', () => {
    const worksheet: PrintWorksheet = {
      getSheetName: () => 'Mixed',
      getSheet: () => ({
        getRowVisible: () => true,
        getColVisible: () => true,
      }),
      getLastRow: () => 0,
      getLastColumn: () => 1,
      getRowHeight: () => 20,
      getColumnWidth: () => 100,
      getMergedRanges: () => [],
      getRange: ((row: number, column: number) => ({
        getDisplayValues: () => [['x', 'y']],
        getValues: () => [[1, 2]],
        getCellStyleData: () => (column === 1 ? { bd: { b: { s: 1 } } } : null),
      })) as PrintWorksheet['getRange'],
    }
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    expect(payload.html).toContain('<tr style="height:16.5pt">')
  })

  it('a borderline sheet now fits the page the model plans', () => {
    // The audit's probe2 shape: enough thin-bordered rows to fill a page
    // exactly under the old model. With the border counted, the simulated
    // pagination and the declared row heights agree, so &N stops counting a
    // phantom trailing page. Rows at 15.75+0.75 = 16.5pt on a 733.68pt
    // printable height: 44 rows = 726pt fit, the 45th tips over.
    const rowCount = 46
    const worksheet: PrintWorksheet = {
      getSheetName: () => 'Tall',
      getSheet: () => ({
        getRowVisible: () => true,
        getColVisible: () => true,
      }),
      getLastRow: () => rowCount - 1,
      getLastColumn: () => 0,
      getRowHeight: () => 20,
      getColumnWidth: () => 100,
      getMergedRanges: () => [],
      getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
        getDisplayValues: () =>
          Array.from({ length: numRows ?? 1 }, (_, r) => [`r${row + r + 1}`]).map((cells) =>
            cells.slice(0, numColumns ?? 1),
          ),
        getValues: () => [],
        getCellStyleData: () => ({ bd: { b: { s: 1 } } }),
      })) as PrintWorksheet['getRange'],
    }
    const payload = buildSheetsPrintPayload(
      [
        { worksheet, printAreas: [], printTitles: null },
        { worksheet: borderedWorksheet(null), printAreas: [], printTitles: null },
      ],
      payloadSetup({ header: { center: '&A' } }),
      'Book.pdf',
      'Tall',
    )
    // 46 rows at 16.5pt: 44 per page -> 2 pages; not 3 with an empty tail.
    expect(payload.sheets!.map((sheet) => sheet.pages)).toEqual([2, 1])
    expect((payload.html.match(/<table>/g) ?? []).length).toBe(2)
  })
})

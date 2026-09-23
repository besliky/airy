/**
 * BUG-1614 planned orphan pages on boundary sheets: the one-line text
 * estimate (fontSize x 1.25 + 2pt padding) dominated the DECLARED row
 * height in printedHeightPt, so a default 15pt row with 11pt text printed
 * as 16.5pt where Excel/LibreOffice lay the same row out at 11.25pt (15pt
 * at 75% scale). A sheet sitting near the page boundary planned 40 rows per
 * page instead of 44: &N counted 3 pages against LO's 2, and every bordered
 * sheet at the boundary grew a stub trailing page (repeated titles + a few
 * rows). Now the declared height wins — the estimate only clamps from
 * below — and a text line taller than its row is clamped INTO the row (the
 * cell's line box pins to the declaration), so the text clips at the row
 * edge like Excel's instead of growing it. Wrap-text rows keep the boosted
 * estimate: Chromium grows them past any declared height (BUG-1214), so
 * the plan must lean tall with the render.
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

/// A worksheet over a caller-supplied text grid (100px = 75pt columns,
/// caller-set row height, every cell thin-bordered on the bottom like the
/// audit's probe2 shape).
function borderedGridWorksheet(
  grid: string[][],
  options: { rowHeight?: number; style?: Record<string, unknown> | null } = {},
): PrintWorksheet {
  const rowHeight = options.rowHeight ?? 20 // 20px = 15pt saved
  return {
    getSheetName: () => 'Probe',
    getSheet: () => ({
      getRowVisible: () => true,
      getColVisible: () => true,
    }),
    getLastRow: () => grid.length - 1,
    getLastColumn: () => Math.max(...grid.map((row) => row.length - 1), 0),
    getRowHeight: () => rowHeight,
    getColumnWidth: () => 100,
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        grid
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getValues: () => [],
      getCellStyleData: () => options.style ?? null,
    })) as PrintWorksheet['getRange'],
  }
}

describe('BUG-1614 declared row height in print layout', () => {
  it('a boundary sheet of explicit 15pt rows plans 2 pages, not 3 (the LO layout)', () => {
    // The audit's probe2 shape at 75% scale: A4 printable height 733.68pt
    // gives a 978.24pt capacity, titles 1:2 repeat on every page, every row
    // thin-bordered. Under the old model each row planned 16.5pt (15pt
    // saved inflated to the 15.75pt text estimate + border): 57 rows per
    // page -> 3 pages, page 3 a stub of repeated titles + 4 rows. LO lays
    // the same sheet out on 2 pages. With the declaration winning, each row
    // plans 15.75pt (15pt saved + 0.75pt border) -> 60 + 58 rows.
    const bodyRows = 118
    const grid = Array.from({ length: bodyRows + 2 }, (_, row) => [`r${row + 1}`])
    const payload = buildSheetsPrintPayload(
      [
        {
          worksheet: borderedGridWorksheet(grid, { style: { bd: { b: { s: 1 } } } }),
          printAreas: [],
          printTitles: '1:2',
        },
        { worksheet: borderedGridWorksheet([['x']]), printAreas: [], printTitles: null },
      ],
      payloadSetup({ scale: 75, header: { center: '&A' } }),
      'Book.pdf',
      'Probe',
    )
    // 2 pages like LibreOffice; the old model planned 3 with a stub tail.
    expect(payload.sheets!.map((sheet) => sheet.pages)).toEqual([2, 1])
    // Rows plan at the declared height + border, never the text estimate.
    expect(payload.html).toContain('<tr style="height:15.75pt">')
    expect(payload.html).not.toContain('<tr style="height:16.5pt">')
    // The 11pt line clamps into the declaration minus padding and border.
    expect(payload.html).toContain('line-height:12.25pt')
  })

  it('a bordered boundary sheet grows no stub trailing page', () => {
    // Same shape one row past the boundary: the old model put 57 + 57 + 1
    // (a titles-only stub) on 3 pages; the new plan fits 60 + 59 on 2.
    const bodyRows = 119
    const grid = Array.from({ length: bodyRows + 2 }, (_, row) => [`r${row + 1}`])
    const payload = buildSheetsPrintPayload(
      [
        {
          worksheet: borderedGridWorksheet(grid, { style: { bd: { b: { s: 1 } } } }),
          printAreas: [],
          printTitles: '1:2',
        },
        { worksheet: borderedGridWorksheet([['x']]), printAreas: [], printTitles: null },
      ],
      payloadSetup({ scale: 75, header: { center: '&A' } }),
      'Book.pdf',
      'Probe',
    )
    expect(payload.sheets!.map((sheet) => sheet.pages)).toEqual([2, 1])
  })

  it('a row with oversized content keeps its content and its saved height', () => {
    // Excel keeps a row at its saved height and clips the oversized text at
    // the row edge; the printout must do the same — the text still prints
    // (nothing lost from the document), the <tr> stays at the saved 15pt,
    // and the line box clamps so the rendered row cannot grow past the
    // height the pagination planned.
    const payload = buildSheetsPrintPayload(
      [
        {
          worksheet: borderedGridWorksheet([['IMPORTANT']], {
            style: { fs: 20 },
          }),
          printAreas: [],
          printTitles: null,
        },
      ],
      payloadSetup(),
      'Book.pdf',
      'Probe',
    )
    expect(payload.html).toContain('>IMPORTANT<')
    expect(payload.html).toContain('<tr style="height:15pt">')
    expect(payload.html).not.toContain('<tr style="height:15.75pt">')
    expect(payload.html).toContain('line-height:13pt')
  })

  it('wrap-text rows keep the text-boosted height the render leans on', () => {
    // Wrap-text cells lay out on several lines and Chromium grows the row
    // past any declared height (BUG-1214 residual), so those rows keep the
    // estimate as their lower-bound plan — and take no line clamp (there is
    // nothing to clamp: the render is taller than one line by design).
    const payload = buildSheetsPrintPayload(
      [
        {
          worksheet: borderedGridWorksheet([['a long wrapped line']], {
            style: { tb: 3 },
          }),
          printAreas: [],
          printTitles: null,
        },
      ],
      payloadSetup(),
      'Book.pdf',
      'Probe',
    )
    expect(payload.html).toContain('<tr style="height:15.75pt">')
    expect(payload.html).not.toContain('line-height:')
  })
})

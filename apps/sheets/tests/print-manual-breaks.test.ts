/**
 * BUG-1500 manual page breaks in print: Excel honours rowBreaks/colBreaks
 * (Page Layout > Breaks, or the file's <rowBreaks>/<colBreaks>) at a fixed
 * print scale; the print layout dropped them, so a sheet with a break after
 * row 40 printed page 1 through row 43. The breaks now flow through
 * EffectivePageSetup (the active sheet's) and per-sheet print jobs into the
 * tile pagination; fit-to-page ignores them like Excel.
 */
import { describe, expect, it } from 'vitest'

import { buildSheetsPrintPayload, type PrintWorksheet } from '../src/renderer/print-html'
import {
  resolveEffectivePageSetup,
  resolveSheetPageBreaks,
  type EffectivePageSetup,
} from '../src/renderer/print-settings'
import type { PageSetupJournalState, StructuralJournalOp } from '../src/renderer/edit-journal'

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

/// A worksheet over a caller-supplied text grid; 100px columns (75pt),
/// caller-set row height (default 20px = 15pt).
function gridWorksheet(
  grid: string[][],
  options: { rowHeight?: number; columnWidth?: number } = {},
): PrintWorksheet {
  const rowHeight = options.rowHeight ?? 20
  const columnWidth = options.columnWidth ?? 100
  return {
    getSheetName: () => 'Grid',
    getSheet: () => ({
      getRowVisible: () => true,
      getColVisible: () => true,
    }),
    getLastRow: () => grid.length - 1,
    getLastColumn: () => Math.max(...grid.map((row) => row.length - 1), 0),
    getRowHeight: () => rowHeight,
    getColumnWidth: () => columnWidth,
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        grid
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getValues: () => [],
      getCellStyleData: () => null,
    })) as PrintWorksheet['getRange'],
  }
}

function tablesOf(html: string): string[] {
  return html.match(/<table>[\s\S]*?<\/table>/g) ?? []
}

function payloadOf(
  worksheet: PrintWorksheet,
  setup: Partial<EffectivePageSetup> = {},
  job: { rowBreaks?: readonly number[]; colBreaks?: readonly number[] } = {},
  pageOrder: 'down-then-over' | 'over-then-down' = 'down-then-over',
) {
  return buildSheetsPrintPayload(
    [{ worksheet, printAreas: [], printTitles: null, ...job }],
    payloadSetup(setup),
    'Book.pdf',
    'Grid',
    undefined,
    pageOrder,
  )
}

describe('buildSheetsPrintPayload manual page breaks', () => {
  it('cuts a tile at a manual row break (down-then-over keeps stripe-major order)', () => {
    // 60 rows at 15pt fit one A4 page height (~733pt); the manual break
    // after row 40 (0-based index 40 = above row 41) must still cut.
    const grid = Array.from({ length: 60 }, (_, row) => [`r${row + 1}`])
    const payload = payloadOf(gridWorksheet(grid), {}, { rowBreaks: [40] })
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(2)
    expect(tables[0]).toContain('>r40<')
    expect(tables[0]).not.toContain('>r41<')
    expect(tables[1]).toContain('>r41<')
    expect(tables[1]).toContain('>r60<')
  })

  it('ignores manual breaks under fit-to-page like Excel', () => {
    const grid = Array.from({ length: 60 }, (_, row) => [`r${row + 1}`])
    const payload = payloadOf(
      gridWorksheet(grid),
      { fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
      { rowBreaks: [40], colBreaks: [0] },
    )
    expect(tablesOf(payload.html)).toHaveLength(1)
  })

  it('breaks each column stripe at the row break, walking stripes outer', () => {
    // Two stripes (12 columns of 75pt on ~495pt) × two row segments: the
    // down-then-over order prints stripe A's pages before stripe B's.
    const grid = Array.from({ length: 60 }, (_, row) =>
      Array.from({ length: 12 }, (_, column) => `r${row + 1}c${column + 1}`),
    )
    const payload = payloadOf(gridWorksheet(grid), {}, { rowBreaks: [40] })
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(4)
    expect(tables[0]).toContain('>r1c1<')
    expect(tables[0]).not.toContain('>r41c1<')
    expect(tables[1]).toContain('>r41c1<')
    expect(tables[2]).toContain('>r1c7<')
    expect(tables[2]).not.toContain('>r1c1<')
    expect(tables[3]).toContain('>r41c7<')
  })

  it('cuts a manual column break as a stripe edge', () => {
    // Six columns of 75pt fit one page; a break above column 4 forces the
    // boundary early.
    const grid = [['c1', 'c2', 'c3', 'c4', 'c5', 'c6']]
    const payload = payloadOf(gridWorksheet(grid), {}, { colBreaks: [3] })
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(2)
    expect(tables[0]).toContain('>c3<')
    expect(tables[0]).not.toContain('>c4<')
    expect(tables[1]).toContain('>c4<')
  })

  it('cuts over-then-down bands at manual breaks on top of capacity', () => {
    // 150pt rows: four per capacity band (~733pt) without breaks; the manual
    // break above row 3 (0-based 2) moves the band edge to row 2 regardless.
    const grid = Array.from({ length: 6 }, (_, row) =>
      Array.from({ length: 12 }, (_, column) => `r${row + 1}c${column + 1}`),
    )
    const payload = payloadOf(
      gridWorksheet(grid, { rowHeight: 200 }),
      {},
      { rowBreaks: [2] },
      'over-then-down',
    )
    const tables = tablesOf(payload.html)
    // Bands [1..2] (manual edge) and [3..6]; × two stripes = four tiles.
    expect(tables).toHaveLength(4)
    expect(tables[0]).toContain('>r2c1<')
    expect(tables[0]).not.toContain('>r3c1<')
    expect(tables[2]).toContain('>r3c1<')
    expect(tables[2]).not.toContain('>r2c1<')
  })

  it('drops breaks at or before the first body row and past the area', () => {
    const grid = Array.from({ length: 30 }, (_, row) => [`r${row + 1}`])
    const noBreaks = payloadOf(
      gridWorksheet(grid),
      {},
      {
        rowBreaks: [0, 99],
      },
    )
    expect(tablesOf(noBreaks.html)).toHaveLength(1)
    // With title rows, a break inside the titles would cut a titles-only
    // segment: dropped too.
    const titled = buildSheetsPrintPayload(
      [
        {
          worksheet: gridWorksheet(grid),
          printAreas: [],
          printTitles: '1:1',
          rowBreaks: [0, 1],
        },
      ],
      payloadSetup(),
      'Book.pdf',
      'Grid',
    )
    expect(tablesOf(titled.html)).toHaveLength(1)
  })

  it('falls back to the setup breaks when the job carries none', () => {
    const grid = Array.from({ length: 60 }, (_, row) => [`r${row + 1}`])
    const payload = buildSheetsPrintPayload(
      [{ worksheet: gridWorksheet(grid), printAreas: [], printTitles: null }],
      payloadSetup({ rowBreaks: [10, 20] }),
      'Book.pdf',
      'Grid',
    )
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(3)
    expect(tables[0]).toContain('>r10<')
    expect(tables[1]).toContain('>r11<')
    expect(tables[2]).toContain('>r21<')
  })

  it('keeps a job break set of [] from leaking the setup set (workbook jobs)', () => {
    const grid = Array.from({ length: 60 }, (_, row) => [`r${row + 1}`])
    const payload = buildSheetsPrintPayload(
      [{ worksheet: gridWorksheet(grid), printAreas: [], printTitles: null, rowBreaks: [] }],
      payloadSetup({ rowBreaks: [40] }),
      'Book.pdf',
      'Grid',
    )
    expect(tablesOf(payload.html)).toHaveLength(1)
  })

  it('repeats title rows on every manual-break segment and counts the pages', () => {
    const grid = Array.from({ length: 60 }, (_, row) => [`r${row + 1}`])
    const payload = buildSheetsPrintPayload(
      [
        { worksheet: gridWorksheet(grid), printAreas: [], printTitles: '1:1', rowBreaks: [40] },
        { worksheet: gridWorksheet([['x']]), printAreas: [], printTitles: null },
      ],
      payloadSetup({ header: { center: '&A' } }),
      'Book.pdf',
      'Grid',
    )
    const tables = tablesOf(payload.html)
    // Two manual-break segments plus the second sheet's single table.
    expect(tables).toHaveLength(3)
    // Both segments' thead repeats the title row; the second segment's body
    // starts at row 41.
    expect((tables[0]!.match(/>r1</g) ?? []).length).toBe(1)
    expect((tables[1]!.match(/>r1</g) ?? []).length).toBe(1)
    expect(tables[1]).toContain('>r41<')
    expect(tables[1]).not.toContain('>r2<')
    // The &A ranged passes plan one page per manual-break segment.
    expect(payload.sheets!.map((sheet) => sheet.pages)).toEqual([2, 1])
  })

  it('honours merge clamps across a manual-break segment edge', () => {
    // A merge anchored above the break spans it: the first segment clamps
    // the rowspan, the second paints styled fillers.
    const grid = [
      ['m', 'x'],
      ['', 'x'],
      ['', 'x'],
    ]
    const worksheet = {
      ...gridWorksheet(grid),
      getMergedRanges: () => [
        { getRow: () => 0, getColumn: () => 0, getWidth: () => 1, getHeight: () => 3 },
      ],
      getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
        getDisplayValues: () =>
          grid
            .slice(row, row + (numRows ?? 1))
            .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
        getValues: () => [],
        getCellStyleData: () => (row === 0 && column === 0 ? { bg: { rgb: '#00ff00' } } : null),
      })) as PrintWorksheet['getRange'],
    }
    const payload = payloadOf(worksheet, {}, { rowBreaks: [2] })
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(2)
    expect(tables[0]).toContain('rowspan="2"')
    // The continuation row paints the anchored fill.
    expect(tables[1]).toMatch(/<td style="[^"]*background:#00ff00/)
  })
})

describe('resolveSheetPageBreaks', () => {
  it('prefers the journal replacement set over the file breaks', () => {
    const journal = { rowBreaks: [5], colBreaks: [] } as PageSetupJournalState
    expect(resolveSheetPageBreaks(journal, { rowBreaks: [2, 9], colBreaks: [1] }, [])).toEqual({
      rowBreaks: [5],
      colBreaks: [],
    })
  })

  it('maps file breaks through the session structural ops', () => {
    // A file break above row 3 (0-based) with 2 rows inserted above it.
    const ops = [{ kind: 'insert-rows', index: 0, count: 2 }] as unknown as StructuralJournalOp[]
    expect(resolveSheetPageBreaks(undefined, { rowBreaks: [3] }, ops)).toEqual({
      rowBreaks: [5],
      colBreaks: [],
    })
  })

  it('drops file breaks deleted by structural ops and negative ids', () => {
    const ops = [{ kind: 'remove-rows', index: 1, count: 3 }] as unknown as StructuralJournalOp[]
    expect(resolveSheetPageBreaks(undefined, { rowBreaks: [2, 4, 0] }, ops)).toEqual({
      rowBreaks: [1],
      colBreaks: [],
    })
  })

  it('flows through resolveEffectivePageSetup onto the setup', () => {
    const setup = resolveEffectivePageSetup({}, null, null, [], { rowBreaks: [7, 12] })
    expect(setup.rowBreaks).toEqual([7, 12])
    expect(setup.colBreaks).toEqual([])
  })
})

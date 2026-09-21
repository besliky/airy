/**
 * BUG-1521 hidden rows/columns in print output: Excel and LibreOffice skip
 * hidden rows and columns on every print path, and a filter is just row
 * hiding — "apply a filter, then export the PDF" must not leak the
 * filtered-out rows. The payload builder reads the sheet's visibility
 * (`getRowVisible` is false for manual hide, outline collapse, AND filter),
 * drops hidden rows/columns from the layout (cells, headings, colgroups,
 * stripes, and the fit-to-page pagination), and keeps a merge visible
 * through its printed part when its anchor sits on a hidden row/column.
 */
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ BrowserWindow: class {}, dialog: {} }))
vi.mock('@airy-office/electron-utils', () => ({ showSaveDialogWithMemory: vi.fn() }))

import { buildPrintRequest, type PageLayoutContext } from '../src/renderer/page-layout-actions'
import { buildSheetsPrintPayload, type PrintWorksheet } from '../src/renderer/print-html'
import type { EffectivePageSetup } from '../src/renderer/print-settings'
import { createEditJournal } from '../src/renderer/edit-journal'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

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

interface WorksheetOptions {
  hiddenRows?: readonly number[]
  hiddenColumns?: readonly number[]
  merges?: { row: number; column: number; width: number; height: number }[]
  rowHeight?: number
  cellStyle?: (row: number, column: number) => Record<string, unknown> | null
}

/// A worksheet over a caller-supplied text grid (100px = 75pt columns,
/// 20px = 15pt rows) whose sheet visibility hides the given row/column
/// indexes — the same `getSheet().getRowVisible`/`getColVisible` reads the
/// live Univer model answers for hidden AND filtered rows/columns.
function gridWorksheet(grid: string[][], options: WorksheetOptions = {}): PrintWorksheet {
  const hiddenRows = new Set(options.hiddenRows ?? [])
  const hiddenColumns = new Set(options.hiddenColumns ?? [])
  return {
    getSheetName: () => 'Grid',
    getLastRow: () => grid.length - 1,
    getLastColumn: () => Math.max(...grid.map((row) => row.length - 1), 0),
    getRowHeight: () => options.rowHeight ?? 20,
    getColumnWidth: () => 100,
    getSheet: () => ({
      getRowVisible: (row: number) => !hiddenRows.has(row),
      getColVisible: (column: number) => !hiddenColumns.has(column),
    }),
    getMergedRanges: () =>
      (options.merges ?? []).map((merge) => ({
        getRow: () => merge.row,
        getColumn: () => merge.column,
        getWidth: () => merge.width,
        getHeight: () => merge.height,
      })),
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        grid
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getValues: () => [],
      getCellStyleData: () => options.cellStyle?.(row, column) ?? null,
    })) as PrintWorksheet['getRange'],
  }
}

function tablesOf(html: string): string[] {
  return html.match(/<table>[\s\S]*?<\/table>/g) ?? []
}

function rowsOf(table: string): string[] {
  return table.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) ?? []
}

describe('hidden rows and columns in print output', () => {
  it('drops hidden rows from the payload (outline probe: a hidden run prints nothing)', () => {
    // probe4's outline sheet: rows 25-30 (indexes 24-29) hidden — Excel/LO
    // print row 24 then row 31; the app printed all six with their text.
    const grid = Array.from({ length: 32 }, (_, row) => [`row${row + 1}`])
    const worksheet = gridWorksheet(grid, { hiddenRows: [24, 25, 26, 27, 28, 29] })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    for (const hidden of ['row25', 'row26', 'row27', 'row28', 'row29', 'row30']) {
      expect(payload.html).not.toContain(hidden)
    }
    expect(payload.html).toContain('row24')
    expect(payload.html).toContain('row31')
    const rows = rowsOf(tablesOf(payload.html)[0] ?? '')
    expect(rows).toHaveLength(26)
  })

  it('drops hidden columns — no cell, no <col>, no heading letter', () => {
    const grid = [
      ['a1', 'b1', 'c1'],
      ['a2', 'b2', 'c2'],
    ]
    const worksheet = gridWorksheet(grid, { hiddenColumns: [1] })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup({ printHeadings: true }),
      'Book.pdf',
      'S',
    )
    const table = tablesOf(payload.html)[0] ?? ''
    expect(table).not.toContain('b1')
    expect(table).not.toContain('b2')
    // Column B's letter is gone from the heading strip and no zero-width
    // <col> keeps its slot: A and C stay in their own slots.
    expect(table).not.toContain('>B</th>')
    expect(table).toContain('>A</th>')
    expect(table).toContain('>C</th>')
    expect((table.match(/<col /g) ?? []).length).toBe(3) // heading strip + A + C
    for (const row of rowsOf(table).slice(1)) {
      expect((row.match(/<td/g) ?? []).length).toBe(2)
    }
    expect(payload.html).toContain('a1')
    expect(payload.html).toContain('c2')
  })

  it('treats filter-hidden rows exactly like hidden rows (one visibility read)', () => {
    // A filter hides rows through the same row-visibility flag a manual
    // hide uses (`getRowVisible` is `!isRowFiltered && rawVisible` in the
    // model), so the filtered export cannot carry the filtered-out text.
    const grid = [['Berlin'], ['Rome'], ['Oslo'], ['Tokyo']]
    const filtered = gridWorksheet(grid, { hiddenRows: [1, 3] })
    const payload = buildSheetsPrintPayload(
      [{ worksheet: filtered, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    expect(payload.html).toContain('Berlin')
    expect(payload.html).toContain('Oslo')
    expect(payload.html).not.toContain('Rome')
    expect(payload.html).not.toContain('Tokyo')
  })

  it('excludes hidden rows inside a selection print job', async () => {
    const { context } = selectionRuntime([1])
    const { request } = await buildPrintRequest(context, { scope: 'selection' })
    expect(request.html).toContain('sel-r1')
    expect(request.html).not.toContain('sel-r2')
    expect(request.html).toContain('sel-r3')
  })

  it('counts only printed rows toward the page structure (hidden rows are free)', () => {
    // 500px (375pt) rows: two per A4 page. Hiding every other row keeps the
    // printed sheet to the visible rows' page count — the hidden ones must
    // not pad the fit-to-page pagination the PDF pages come from.
    const grid = Array.from({ length: 6 }, (_, row) => [`r${row + 1}`])
    const worksheet = gridWorksheet(grid, { rowHeight: 500, hiddenRows: [1, 3, 5] })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup({ fitToPage: true, fitToHeight: 0, fitToWidth: 1 }),
      'Book.pdf',
      'S',
    )
    const rows = rowsOf(tablesOf(payload.html)[0] ?? '')
    expect(rows).toHaveLength(3)
  })

  it('keeps a merge visible through its printed part when the anchor row is hidden', () => {
    // A1:A3 merged, row 1 hidden: Excel still shows the merged cell across
    // the visible rows 2-3. The printed anchor relocates to the first
    // printed row, keeping the merge's text and fill.
    const grid = [
      ['merged', ''],
      ['x', 'y'],
      ['z', 'w'],
    ]
    const worksheet = gridWorksheet(grid, {
      hiddenRows: [0],
      merges: [{ row: 0, column: 0, width: 1, height: 3 }],
      cellStyle: (row, column) => (row === 0 && column === 0 ? { bg: { rgb: '#ff8c00' } } : null),
    })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup({ printGridlines: true }),
      'Book.pdf',
      'S',
    )
    const table = tablesOf(payload.html)[0] ?? ''
    // The relocated anchor prints the merge's text over the two printed
    // rows, in the anchor's fill.
    expect(table).toMatch(/<td rowspan="2"[^>]*>merged<\/td>/)
    expect(table).toMatch(/<td rowspan="2" style="[^"]*background:#ff8c00/)
    // The covered cell in the second printed row is consumed (no filler
    // shifts 'y' or 'w' out of its column slot).
    const rows = rowsOf(table)
    expect(rows).toHaveLength(2)
    expect((rows[0] ?? '').match(/<td/g)).toHaveLength(2)
    expect(rows[0]).toContain('y')
    // The covered cell under the merge is consumed: 'w' keeps its column.
    expect((rows[1] ?? '').match(/<td/g)).toHaveLength(1)
    expect(rows[1]).toContain('w')
    expect(table).not.toContain('>z<')
  })

  it('relocates a merge whose anchor column is hidden, not just the anchor row', () => {
    // A1:C1 merged, column A hidden: the merge still spans B1:C1 visually.
    const grid = [
      ['merged', '', ''],
      ['a2', 'b2', 'c2'],
    ]
    const worksheet = gridWorksheet(grid, {
      hiddenColumns: [0],
      merges: [{ row: 0, column: 0, width: 3, height: 1 }],
    })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    const table = tablesOf(payload.html)[0] ?? ''
    expect(table).toMatch(/<td colspan="2"[^>]*>merged<\/td>/)
    // Row 2 keeps its two printed cells in their slots.
    const rows = rowsOf(table)
    expect(rows[1]).toContain('b2')
    expect(rows[1]).toContain('c2')
    expect((rows[1] ?? '').match(/<td/g)).toHaveLength(2)
  })

  it('folds hidden middle rows out of a visible anchor merge span', () => {
    // A1:A4 merged, rows 2-3 hidden: the merge spans the printed rows only
    // (a raw rowspan of 4 would reach past the emitted <tr>s).
    const grid = [
      ['m', 'a'],
      ['', 'b'],
      ['', 'c'],
      ['', 'd'],
    ]
    const worksheet = gridWorksheet(grid, {
      hiddenRows: [1, 2],
      merges: [{ row: 0, column: 0, width: 1, height: 4 }],
    })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    const table = tablesOf(payload.html)[0] ?? ''
    expect(table).toMatch(/<td rowspan="2"[^>]*>m<\/td>/)
    const rows = rowsOf(table)
    // Printed rows 1 and 4: each keeps one cell beside the merge.
    expect(rows).toHaveLength(2)
    expect((rows[1] ?? '').match(/<td/g)).toHaveLength(1)
    expect(rows[1]).toContain('d')
  })

  it('drops a merge that has no printed cell at all', () => {
    // A2:B2 merged inside two hidden rows: nothing of it prints, and the
    // surrounding columns must not keep phantom covered slots.
    const grid = [
      ['a1', 'b1'],
      ['m1', 'm2'],
      ['a3', 'b3'],
    ]
    const worksheet = gridWorksheet(grid, {
      hiddenRows: [1],
      merges: [{ row: 1, column: 0, width: 2, height: 1 }],
    })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    const table = tablesOf(payload.html)[0] ?? ''
    expect(table).not.toContain('m1')
    expect(table).not.toContain('m2')
    for (const row of rowsOf(table)) {
      expect((row.match(/<td/g) ?? []).length).toBe(2)
    }
  })

  it('still repeats title rows over hidden body rows (titles print as saved)', () => {
    const grid = [
      ['title', 't2'],
      ['body1', 'b1'],
      ['body2', 'b2'],
    ]
    const worksheet = gridWorksheet(grid, { hiddenRows: [2] })
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: ['A2:C3'], printTitles: '1:1' }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    const table = tablesOf(payload.html)[0] ?? ''
    expect(table).toContain('title')
    expect(table).toContain('body1')
    expect(table).not.toContain('body2')
  })
})

/// A page-layout context whose active sheet prints a selection (B1:B3)
/// with caller-hidden rows — the selection override path of the dialog.
function selectionRuntime(hiddenRows: readonly number[]): { context: PageLayoutContext } {
  const grid = [
    ['a1', 'sel-r1'],
    ['a2', 'sel-r2'],
    ['a3', 'sel-r3'],
  ]
  const worksheet = {
    ...gridWorksheet(grid, { hiddenRows }),
    getSheetId: () => 'sheet-1',
  }
  const workbook = {
    getId: () => 'wb-1',
    getActiveSheet: () => worksheet,
    getActiveRange: () => ({
      getRow: () => 0,
      getColumn: () => 1,
      getWidth: () => 1,
      getHeight: () => 3,
    }),
    getSheets: () => [worksheet],
  }
  const state = {
    file: { sheets: [], name: 'Book.xlsx', sessionId: 'session-1' },
    editJournal: createEditJournal(),
    sheetFilePageSetups: new Map(),
    flags: { preloadComplete: true, preloadRunning: false },
  } as unknown as LazyWorkbookState
  return {
    context: {
      univerRef: {
        current: { univerAPI: { getActiveWorkbook: () => workbook } } as unknown as UniverRuntime,
      },
      lazyWorkbookRef: { current: state },
      setMessage: () => {},
    } as unknown as PageLayoutContext,
  }
}

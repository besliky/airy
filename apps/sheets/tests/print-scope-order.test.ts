/**
 * PAR-207 print scope and page order: the Print dialog's job builder picks
 * the sheets by scope (selection override, active sheet, entire workbook
 * over visible sheets), hands collate through to the print job, and the
 * layout tiles wide sheets into column stripes ordered down-then-over or
 * over-then-down like Excel.
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

/// A worksheet over a caller-supplied text grid; 100px columns (75pt),
/// caller-set row height (default 20px = 15pt).
function gridWorksheet(
  grid: string[][],
  options: { rowHeight?: number; columnWidth?: number; lastRow?: number; lastColumn?: number } = {},
): PrintWorksheet {
  const rowHeight = options.rowHeight ?? 20
  const columnWidth = options.columnWidth ?? 100
  return {
    getSheetName: () => 'Grid',
    getLastRow: () => options.lastRow ?? grid.length - 1,
    getLastColumn: () => options.lastColumn ?? Math.max(...grid.map((row) => row.length - 1), 0),
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

describe('buildSheetsPrintPayload', () => {
  it('prints every visible sheet of a workbook job, each starting a page', () => {
    const first = gridWorksheet([['a1']])
    const second = gridWorksheet([['b1', 'b2']])
    const payload = buildSheetsPrintPayload(
      [
        { worksheet: first, printAreas: [], printTitles: null },
        { worksheet: second, printAreas: [], printTitles: null },
      ],
      payloadSetup(),
      'Book.pdf',
      'First',
    )
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(2)
    expect(tables[0]).toContain('a1')
    expect(tables[1]).toContain('b2')
    expect(payload.html).toContain('table + table { break-before: page; }')
  })

  it('skips blank sheets under skipWhenEmpty but keeps explicit print areas', () => {
    const blank = gridWorksheet([['']], { lastRow: -1, lastColumn: -1 })
    const withArea = gridWorksheet([['x']])
    const payload = buildSheetsPrintPayload(
      [
        { worksheet: blank, printAreas: [], printTitles: null, skipWhenEmpty: true },
        { worksheet: withArea, printAreas: ['A1:A1'], printTitles: null, skipWhenEmpty: true },
      ],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    expect(tablesOf(payload.html)).toHaveLength(1)
    expect(payload.html).toContain('x')
  })

  it('tiles a wide sheet into column stripes without page order (down, then over)', () => {
    // A4 portrait printable width is ~494.6pt; 75pt columns → 6 per stripe.
    const grid = [['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8']]
    const payload = buildSheetsPrintPayload(
      [{ worksheet: gridWorksheet(grid), printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(2)
    expect(tables[0]).toContain('c6')
    expect(tables[0]).not.toContain('c7')
    expect(tables[1]).toContain('c7')
    expect(tables[1]).toContain('c8')
  })

  it('orders tiles across first when page order is over-then-down', () => {
    // 375pt rows → one row per band; 8 columns → two stripes: four tiles.
    const grid = [
      ['r1c1', 'r1c2', 'r1c3', 'r1c4', 'r1c5', 'r1c6', 'r1c7', 'r1c8'],
      ['r2c1', 'r2c2', 'r2c3', 'r2c4', 'r2c5', 'r2c6', 'r2c7', 'r2c8'],
    ]
    const payload = buildSheetsPrintPayload(
      [{ worksheet: gridWorksheet(grid, { rowHeight: 500 }), printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
      undefined,
      'over-then-down',
    )
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(4)
    // Tile order: (row 1 × left stripe), (row 1 × right stripe), then row 2.
    expect(tables[0]).toContain('r1c1')
    expect(tables[0]).not.toContain('r1c7')
    expect(tables[1]).toContain('r1c7')
    expect(tables[1]).not.toContain('r1c1')
    expect(tables[2]).toContain('r2c1')
    expect(tables[3]).toContain('r2c8')
  })

  it('resolves &A per sheet on workbook jobs (per-sheet template sets)', () => {
    const alpha = { ...gridWorksheet([['a1']]), getSheetName: () => 'Alpha' }
    const beta = { ...gridWorksheet([['b1']]), getSheetName: () => 'Beta' }
    const jobs = [
      { worksheet: alpha, printAreas: [] as string[], printTitles: null },
      { worksheet: beta, printAreas: [] as string[], printTitles: null },
    ]
    const payload = buildSheetsPrintPayload(
      jobs,
      payloadSetup({ header: { center: '&A' } }),
      'Book.pdf',
      'Alpha',
    )
    // each sheet carries its own template set: &A names the page's owner
    expect(payload.sheets).toHaveLength(2)
    expect(payload.sheets![0]!.headerTemplate).toContain('Alpha')
    expect(payload.sheets![0]!.headerTemplate).not.toContain('Beta')
    expect(payload.sheets![1]!.headerTemplate).toContain('Beta')
    expect(payload.sheets![1]!.headerTemplate).not.toContain('Alpha')
    expect(payload.sheets!.map((s) => s.pages)).toEqual([1, 1])
    // the job-level templates stay the active sheet's (single-pass jobs)
    expect(payload.headerTemplate).toContain('Alpha')

    // variants resolve per sheet too (differentFirst / differentOddEven)
    const varied = buildSheetsPrintPayload(
      jobs,
      payloadSetup({
        header: { center: '&A' },
        firstPage: { header: { left: 'First &A' }, footer: null },
        evenPages: { header: { right: 'Even &A' }, footer: null },
      }),
      'Book.pdf',
      'Alpha',
    )
    expect(varied.sheets![1]!.firstPage!.headerTemplate).toContain('Beta')
    expect(varied.sheets![1]!.evenPages!.headerTemplate).toContain('Beta')
    expect(varied.sheets![0]!.firstPage!.headerTemplate).toContain('Alpha')
  })

  it("counts each sheet's pages from the tile pagination", () => {
    // 2 rows at 375pt on a 733.68pt printable height = 2 pages; then 1 row
    const tall = {
      ...gridWorksheet([['a'], ['b']], { rowHeight: 500 }),
      getSheetName: () => 'Tall',
    }
    const short = { ...gridWorksheet([['c']]), getSheetName: () => 'Short' }
    const payload = buildSheetsPrintPayload(
      [
        { worksheet: tall, printAreas: [], printTitles: null },
        { worksheet: short, printAreas: [], printTitles: null },
      ],
      payloadSetup({ header: { center: '&A' } }),
      'Book.pdf',
      'Tall',
    )
    expect(payload.sheets!.map((s) => s.pages)).toEqual([2, 1])
  })

  it('omits the per-sheet sets without &A or with a single sheet', () => {
    const alpha = { ...gridWorksheet([['a1']]), getSheetName: () => 'Alpha' }
    const beta = { ...gridWorksheet([['b1']]), getSheetName: () => 'Beta' }
    const noCode = buildSheetsPrintPayload(
      [
        { worksheet: alpha, printAreas: [], printTitles: null },
        { worksheet: beta, printAreas: [], printTitles: null },
      ],
      payloadSetup({ header: { center: 'Page &P of &N' } }),
      'Book.pdf',
      'Alpha',
    )
    expect(noCode.sheets).toBeUndefined()

    const single = buildSheetsPrintPayload(
      [{ worksheet: alpha, printAreas: [], printTitles: null }],
      payloadSetup({ header: { center: '&A' } }),
      'Book.pdf',
      'Alpha',
    )
    expect(single.sheets).toBeUndefined()
    expect(single.headerTemplate).toContain('Alpha')
  })

  it('fills merge continuation cells so nothing shifts into the wrong stripe', () => {
    // A merged anchor in the first stripe spanning into the second: the
    // second stripe's covered cells render empty instead of collapsing.
    const grid = [
      ['m', '', '', '', '', '', '', ''],
      ['m2', 'x', '', '', '', '', '', ''],
    ]
    const worksheet = {
      ...gridWorksheet(grid),
      getMergedRanges: () => [
        { getRow: () => 0, getColumn: () => 0, getWidth: () => 7, getHeight: () => 1 },
      ],
    }
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: [], printTitles: null }],
      payloadSetup(),
      'Book.pdf',
      'S',
    )
    const tables = tablesOf(payload.html)
    expect(tables).toHaveLength(2)
    // The anchor clamps to the six columns of the first stripe.
    expect(/<td colspan="6"[^>]*>m<\/td>/.exec(tables[0] ?? '')).not.toBeNull()
    // The second stripe's first row carries empty filler cells (the merge
    // continues from outside the stripe), and 'x' keeps its column slot.
    const secondRows = (tables[1] ?? '').match(/<tr[^>]*>[\s\S]*?<\/tr>/g) ?? []
    expect(secondRows[0]).not.toMatch(/<td[^>]*>m/)
    expect((secondRows[0] ?? '').match(/<td/g)).toHaveLength(2)
    const firstRows = (tables[0] ?? '').match(/<tr[^>]*>[\s\S]*?<\/tr>/g) ?? []
    expect(firstRows[1]).toContain('m2')
    expect(firstRows[1]).toContain('x')
  })
})

/// A runtime with two visible sheets plus one hidden, used by the scope tests.
function scopeRuntime() {
  const active = {
    ...gridWorksheet([
      ['a1', 'a2'],
      ['a3', 'a4'],
    ]),
    getSheetId: () => 'sheet-active',
    getSheetName: () => 'Active',
    getZoom: () => 1,
    getConfig: () => ({ hidden: 0 }),
  }
  const other = {
    ...gridWorksheet([['other1']]),
    getSheetId: () => 'sheet-other',
    getSheetName: () => 'Other',
    getZoom: () => 1,
    getConfig: () => ({ hidden: 0 }),
  }
  const hidden = {
    ...gridWorksheet([['secret']]),
    getSheetId: () => 'sheet-hidden',
    getSheetName: () => 'Hidden',
    getZoom: () => 1,
    getConfig: () => ({ hidden: 1 }),
  }
  const workbook = {
    getId: () => 'wb-1',
    getActiveSheet: () => active,
    getActiveRange: () => ({
      getRow: () => 0,
      getColumn: () => 1,
      getWidth: () => 1,
      getHeight: () => 2,
    }),
    getSheets: () => [active, hidden, other],
  }
  return {
    univerAPI: { getActiveWorkbook: () => workbook },
  } as unknown as UniverRuntime
}

function scopeContext(): PageLayoutContext {
  const editJournal = createEditJournal()
  const state = {
    file: { sheets: [], name: 'Book.xlsx', sessionId: 'session-1' },
    editJournal,
    sheetFilePageSetups: new Map(),
    flags: { preloadComplete: true, preloadRunning: false },
  } as unknown as LazyWorkbookState
  return {
    univerRef: { current: scopeRuntime() },
    lazyWorkbookRef: { current: state },
    setMessage: () => {},
  } as unknown as PageLayoutContext
}

describe('buildPrintRequest scopes', () => {
  it('prints the selection as the job print area', async () => {
    const { request } = await buildPrintRequest(scopeContext(), { scope: 'selection' })
    // The active range was column 2 (B), rows 1-2.
    expect(request.html).toContain('a2')
    expect(request.html).toContain('a4')
    expect(request.html).not.toContain('a1')
    expect(request.html).not.toContain('other1')
  })

  it('prints every visible sheet of the workbook, skipping hidden ones', async () => {
    const { request } = await buildPrintRequest(scopeContext(), { scope: 'workbook' })
    expect(request.html).toContain('a1')
    expect(request.html).toContain('other1')
    expect(request.html).not.toContain('secret')
    expect((request.html.match(/<table>/g) ?? []).length).toBe(2)
  })

  it('defaults to the active sheet and forwards collate to the job', async () => {
    const plain = await buildPrintRequest(scopeContext())
    expect(plain.request.html).toContain('a1')
    expect(plain.request.html).not.toContain('other1')
    expect(plain.request.collate).toBeUndefined()
    const collated = await buildPrintRequest(scopeContext(), { collate: true })
    expect(collated.request.collate).toBe(true)
  })
})

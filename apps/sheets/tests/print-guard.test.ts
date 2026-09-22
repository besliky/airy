/**
 * BUG-1603 print guard: a sheet saved at a fixed 100% scale tiles "however
 * it wants" into horizontal strips when it is wider than the paper (UR-04:
 * 14 columns on Letter → 3 strips, 48 pages), and paper inherited from the
 * file (Letter) can differ from the user's locale (A4). The guard detects
 * both pains and the dialog turns them into one-click job overrides
 * (fit-to-width, paper switch) that never touch the saved file — Excel also
 * fits a sheet only when the user asks.
 */
import { describe, expect, it } from 'vitest'

import {
  detectPrintPain,
  localePaperSize,
  PAPER_A4,
  PAPER_LETTER,
} from '../src/renderer/print-guard'
import {
  buildSheetsPrintPayload,
  stripsAcrossJob,
  type PrintSheetJob,
  type PrintWorksheet,
} from '../src/renderer/print-html'
import {
  buildActiveSheetPrintRequest,
  type PageLayoutContext,
} from '../src/renderer/page-layout-actions'
import { createEditJournal } from '../src/renderer/edit-journal'
import type { EffectivePageSetup } from '../src/renderer/print-settings'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

describe('localePaperSize', () => {
  it('maps US-oriented English to Letter and everything else to A4', () => {
    expect(localePaperSize('en')).toBe(PAPER_LETTER)
    expect(localePaperSize('ru')).toBe(PAPER_A4)
    expect(localePaperSize('de')).toBe(PAPER_A4)
    expect(localePaperSize('zh')).toBe(PAPER_A4)
    expect(localePaperSize('ja')).toBe(PAPER_A4)
  })
})

describe('detectPrintPain', () => {
  const base = {
    fitToPage: false,
    stripsAcross: 1,
    paperSize: PAPER_A4,
    localePaperSize: PAPER_A4,
  }

  it('suggests fit-to-width only when a fixed scale tiles across strips', () => {
    expect(detectPrintPain({ ...base, stripsAcross: 3 }).suggestFitToWidth).toBe(true)
    expect(detectPrintPain({ ...base, stripsAcross: 2 }).suggestFitToWidth).toBe(true)
    expect(detectPrintPain({ ...base, stripsAcross: 1 }).suggestFitToWidth).toBe(false)
  })

  it('never suggests fit-to-width when fit-to-page is already on', () => {
    expect(
      detectPrintPain({ ...base, fitToPage: true, stripsAcross: null }).suggestFitToWidth,
    ).toBe(false)
  })

  it('carries the measured strip count through to the hint', () => {
    const guard = detectPrintPain({ ...base, stripsAcross: 3 })
    expect(guard.stripsAcross).toBe(3)
    expect(guard.suggestPaper).toBe(false)
  })

  it('suggests the locale paper only when the job paper differs from it', () => {
    expect(
      detectPrintPain({ ...base, paperSize: PAPER_LETTER, localePaperSize: PAPER_A4 }).suggestPaper,
    ).toBe(true)
    expect(
      detectPrintPain({ ...base, paperSize: PAPER_A4, localePaperSize: PAPER_A4 }).suggestPaper,
    ).toBe(false)
    expect(
      detectPrintPain({ ...base, paperSize: PAPER_A4, localePaperSize: PAPER_LETTER }).suggestPaper,
    ).toBe(true)
  })
})

/// A4 portrait with normal margins unless overridden.
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

/// A fixed-uniform-width worksheet: 100px columns (75pt) by default, 20px
/// rows (15pt). UR-04's shape — many columns (14 × 75pt) over a portrait
/// page (~495pt on A4, ~511pt on Letter) tiles into 3 strips at 100%.
function wideWorksheet(
  columns: number,
  options: { hidden?: readonly number[]; widthPx?: number } = {},
): PrintWorksheet {
  const hidden = new Set(options.hidden ?? [])
  const columnWidthPx = options.widthPx ?? 100
  const row = Array.from({ length: columns }, (_, column) => `c${column + 1}`)
  return {
    getSheetName: () => 'Sheet',
    getLastRow: () => 2,
    getLastColumn: () => columns - 1,
    getRowHeight: () => 20,
    getColumnWidth: () => columnWidthPx,
    getSheet: () => ({
      getRowVisible: () => true,
      getColVisible: (column: number) => !hidden.has(column),
    }),
    getMergedRanges: () => [],
    getRange: ((rowIndex: number, columnIndex: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        Array.from({ length: numRows ?? 1 }, () =>
          Array.from({ length: numColumns ?? 1 }, () => row[columnIndex] ?? ''),
        ),
      getValues: () => [],
      getCellStyleData: () => null,
    })) as PrintWorksheet['getRange'],
  }
}

function jobOf(worksheet: PrintWorksheet): PrintSheetJob {
  return { worksheet, printAreas: [], printTitles: null }
}

describe('stripsAcrossJob (the stripe model at a fixed scale)', () => {
  it('counts 3 strips for 14 plain columns on Letter portrait at 100% (UR-04)', () => {
    expect(
      stripsAcrossJob([jobOf(wideWorksheet(14))], payloadSetup({ paperSize: PAPER_LETTER })),
    ).toBe(3)
  })

  it('counts 3 strips on A4 portrait and fewer on landscape', () => {
    expect(stripsAcrossJob([jobOf(wideWorksheet(14))], payloadSetup())).toBe(3)
    expect(
      stripsAcrossJob([jobOf(wideWorksheet(14))], payloadSetup({ orientation: 'landscape' })),
    ).toBe(2)
  })

  it('hidden columns print no width and can drop a strip', () => {
    expect(stripsAcrossJob([jobOf(wideWorksheet(14, { hidden: [12, 13] }))], payloadSetup())).toBe(
      2,
    )
  })

  it('the heading strip shares the page width with the columns', () => {
    // 8 columns of 60pt fit ~495pt alone, but the 24pt heading strip pushes
    // the last column onto a second strip.
    const setup = payloadSetup({ printHeadings: true })
    expect(
      stripsAcrossJob([jobOf(wideWorksheet(8, { widthPx: 80 }))], {
        ...setup,
        printHeadings: false,
      }),
    ).toBe(1)
    expect(stripsAcrossJob([jobOf(wideWorksheet(8, { widthPx: 80 }))], setup)).toBe(2)
  })

  it('measures the print areas, not the used range', () => {
    const job = { worksheet: wideWorksheet(14), printAreas: ['A1:B10'], printTitles: null }
    expect(stripsAcrossJob([job], payloadSetup())).toBe(1)
  })
})

describe('the fit-to-width override collapses the strips (payload level)', () => {
  const tablesOf = (html: string): number => html.match(/<table>/g)?.length ?? 0

  it('100% tiles into 3 strip tables; fit-to-width prints one strip', () => {
    const wide = wideWorksheet(14)
    const at100 = buildSheetsPrintPayload([jobOf(wide)], payloadSetup(), 'Book.pdf', 'Sheet')
    expect(tablesOf(at100.html)).toBe(3)
    expect(at100.scale).toBe(1)
    const fitWidth = buildSheetsPrintPayload(
      [jobOf(wide)],
      payloadSetup({ fitToPage: true, fitToWidth: 1, fitToHeight: 0 }),
      'Book.pdf',
      'Sheet',
    )
    expect(tablesOf(fitWidth.html)).toBe(1)
    expect(fitWidth.scale).toBeLessThan(1)
  })
})

/// A runtime whose active sheet is the UR-04-like wide sheet.
function wideRuntime(): UniverRuntime {
  const sheet = {
    ...wideWorksheet(14),
    getSheetId: () => 'sheet-1',
    getSheetName: () => 'Worksheet',
    getZoom: () => 1,
    getConfig: () => ({ hidden: 0 }),
  }
  const workbook = {
    getId: () => 'wb-1',
    getActiveSheet: () => sheet,
    getSheets: () => [sheet],
  }
  return { univerAPI: { getActiveWorkbook: () => workbook } } as unknown as UniverRuntime
}

/// The dialog context with the file's saved page setup (Letter, like UR-04).
function wideContext(filePaperSize: number): PageLayoutContext {
  const state = {
    file: { sheets: [], name: 'Book.xlsx', sessionId: 'session-1' },
    editJournal: createEditJournal(),
    sheetFilePageSetups: new Map([['sheet-1', { paperSize: filePaperSize }]]),
    flags: { preloadComplete: true, preloadRunning: false },
  } as unknown as LazyWorkbookState
  return {
    univerRef: { current: wideRuntime() },
    lazyWorkbookRef: { current: state },
    setMessage: () => {},
  } as unknown as PageLayoutContext
}

describe('buildPrintRequest exposes the guard', () => {
  it('flags 3 strips and the foreign Letter paper for the UR-04 shape', async () => {
    const { guard, request } = await buildActiveSheetPrintRequest(wideContext(PAPER_LETTER))
    // Tests run under the zh UI locale, which expects A4.
    expect(guard.suggestFitToWidth).toBe(true)
    expect(guard.stripsAcross).toBe(3)
    expect(guard.suggestPaper).toBe(true)
    expect(guard.localePaperSize).toBe(PAPER_A4)
    expect(request.pageSize).toBe('Letter')
  })

  it('the A4 paper file only flags the strips', async () => {
    const { guard } = await buildActiveSheetPrintRequest(wideContext(PAPER_A4))
    expect(guard.suggestFitToWidth).toBe(true)
    expect(guard.suggestPaper).toBe(false)
  })

  it('applying fit-to-width (1 wide, height free) clears the strip pain', async () => {
    const { guard, request } = await buildActiveSheetPrintRequest(wideContext(PAPER_LETTER), {
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    })
    expect(guard.suggestFitToWidth).toBe(false)
    expect(guard.stripsAcross).toBeNull()
    expect(request.scale).toBeLessThan(1)
    expect(request.html.match(/<table>/g)?.length).toBe(1)
  })

  it('switching the job paper clears the paper pain without touching the file', async () => {
    const { guard, request } = await buildActiveSheetPrintRequest(wideContext(PAPER_LETTER), {
      paperSize: PAPER_A4,
    })
    expect(guard.suggestPaper).toBe(false)
    expect(request.pageSize).toBe('A4')
  })
})

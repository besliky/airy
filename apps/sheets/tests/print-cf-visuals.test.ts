/**
 * BUG-1502 conditional formatting in print: the layout paints CF fills,
 * font colors and data bars from the worksheet's composed CF style — the
 * static getCellStyleData read cannot see them, so PDFs lost every CF
 * visual the grid shows. Icon sets stay deferred (canvas bitmaps).
 */
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ BrowserWindow: class {}, dialog: {} }))
vi.mock('@airy-office/electron-utils', () => ({ showSaveDialogWithMemory: vi.fn() }))

import { buildPrintRequest, type PageLayoutContext } from '../src/renderer/page-layout-actions'
import {
  buildSheetsPrintPayload,
  type PrintConditionalFormatStyle,
  type PrintWorksheet,
} from '../src/renderer/print-html'
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

/// A worksheet over a text grid with per-cell static styles and conditional
/// formatting results.
function cfWorksheet(
  grid: string[][],
  options: {
    style?: (row: number, column: number) => Record<string, unknown> | null
    cf?: (row: number, column: number) => PrintConditionalFormatStyle | null
  } = {},
): PrintWorksheet {
  return {
    getSheetName: () => 'CF',
    getSheet: () => ({
      getRowVisible: () => true,
      getColVisible: () => true,
    }),
    getLastRow: () => grid.length - 1,
    getLastColumn: () => Math.max(...grid.map((row) => row.length - 1), 0),
    getRowHeight: () => 20,
    getColumnWidth: () => 100,
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        grid
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getValues: () => [],
      getCellStyleData: () => options.style?.(row, column) ?? null,
    })) as PrintWorksheet['getRange'],
    getConditionalFormatStyle: (row: number, column: number) => options.cf?.(row, column) ?? null,
  }
}

function payloadOf(
  worksheet: PrintWorksheet,
  setup: Partial<EffectivePageSetup> = {},
): { html: string; tables: string[] } {
  const payload = buildSheetsPrintPayload(
    [{ worksheet, printAreas: [], printTitles: null }],
    payloadSetup(setup),
    'Book.pdf',
    'CF',
  )
  return { html: payload.html, tables: payload.html.match(/<table>[\s\S]*?<\/table>/g) ?? [] }
}

describe('buildSheetsPrintPayload conditional formatting', () => {
  it('paints CF fills and font colors on unstyled cells', () => {
    const { html } = payloadOf(
      cfWorksheet([['low', 'high']], {
        cf: (row, column) =>
          column === 0
            ? { style: { bg: { rgb: '#ffc7ce' }, cl: { rgb: '#9c0006' } } }
            : { style: { bg: { rgb: '#c6efce' }, cl: { rgb: '#006100' } } },
      }),
    )
    expect(html).toContain('background:#ffc7ce')
    expect(html).toContain('color:#9c0006')
    expect(html).toContain('background:#c6efce')
    expect(html).toContain('color:#006100')
  })

  it('layers the CF dxf over the static style (CF fill and font win)', () => {
    const { html } = payloadOf(
      cfWorksheet([['x']], {
        style: () => ({ bg: { rgb: '#ffffff' }, cl: { rgb: '#0000ff' }, bl: 1 }),
        cf: () => ({ style: { bg: { rgb: '#ffeb9c' }, cl: { rgb: '#9c6500' } } }),
      }),
    )
    expect(html).toContain('background:#ffeb9c')
    expect(html).toContain('color:#9c6500')
    expect(html).not.toContain('#ffffff')
    expect(html).not.toContain('#0000ff')
    // The dxf leaves bold undefined: the static bold stays.
    expect(html).toContain('font-weight:700')
  })

  it('keeps the static formatting where the dxf is silent', () => {
    const { html } = payloadOf(
      cfWorksheet([['x']], {
        style: () => ({ it: 1, cl: { rgb: '#123456' } }),
        cf: () => ({ style: { bg: { rgb: '#ddebf7' } } }),
      }),
    )
    expect(html).toContain('font-style:italic')
    expect(html).toContain('color:#123456')
    expect(html).toContain('background:#ddebf7')
  })

  it('prints a data bar as a gradient growing from its axis', () => {
    const { html } = payloadOf(
      cfWorksheet([['v']], {
        cf: () => ({
          dataBar: { color: '#638ec6', value: 50, startPoint: 0, isGradient: false },
          showValue: true,
        }),
      }),
    )
    // Bar covers the left half with hard stops.
    expect(html).toContain(
      'background-image:linear-gradient(90deg,rgba(0,0,0,0) 0%,#638ec6 0%,#638ec6 50%,rgba(0,0,0,0) 50%)',
    )
    expect(html).toContain('>v<')
  })

  it('fades gradient bars to white and grows negative bars left of the axis', () => {
    const { html } = payloadOf(
      cfWorksheet([['pos', 'neg']], {
        cf: (row, column) =>
          column === 0
            ? { dataBar: { color: '#ffbe38', value: 40, startPoint: 0, isGradient: true } }
            : { dataBar: { color: '#ffbe38', value: -50, startPoint: 50, isGradient: true } },
      }),
    )
    expect(html).toContain('#ffbe38 0%,#fff 40%')
    // Negative: white at the left edge, color at the 50% axis.
    expect(html).toContain('#fff 25%,#ffbe38 50%')
  })

  it('hides the value of a bar-only rule but keeps the bar and row height', () => {
    const { tables } = payloadOf(
      cfWorksheet([['secret']], {
        cf: () => ({
          dataBar: { color: '#63be7b', value: 80, startPoint: 0, isGradient: false },
          showValue: false,
        }),
      }),
    )
    expect(tables[0]).not.toContain('secret')
    expect(tables[0]).toContain('#63be7b')
    // The row still declares its saved 15pt height.
    expect(tables[0]).toContain('<tr style="height:15pt">')
  })

  it('prints CF colors of merge fillers anchored above the print area', () => {
    const grid = [
      ['m', '', ''],
      ['', '', ''],
      ['', '', ''],
    ]
    const worksheet = {
      ...cfWorksheet(grid, {
        cf: (row) => (row === 0 ? { style: { bg: { rgb: '#f8696b' } } } : null),
      }),
      getMergedRanges: () => [
        { getRow: () => 0, getColumn: () => 0, getWidth: () => 3, getHeight: () => 3 },
      ],
    }
    const payload = buildSheetsPrintPayload(
      [{ worksheet, printAreas: ['A3:C3'], printTitles: '2:2' }],
      payloadSetup(),
      'Book.pdf',
      'CF',
    )
    const rows = payload.html.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) ?? []
    // The body row's three fillers carry the anchored CF fill.
    expect(rows[1]).toMatch(/<td style="[^"]*background:#f8696b/)
  })
})

/// The runtime wiring: a fake injector serving a fake CF service, with the
/// worksheet built as a class instance (prototype methods) to prove the
/// Object.create delegation used in page-layout-actions.
describe('withConditionalFormatStyle wiring', () => {
  class FakeSheet {
    getSheetId(): string {
      return 'sheet-1'
    }
    getSheetName(): string {
      return 'Wired'
    }
    getLastRow(): number {
      return 0
    }
    getLastColumn(): number {
      return 0
    }
    getRowHeight(): number {
      return 20
    }
    getColumnWidth(): number {
      return 100
    }
    getMergedRanges(): unknown[] {
      return []
    }
    getSheet(): { getRowVisible(row: number): boolean; getColVisible(column: number): boolean } {
      return { getRowVisible: () => true, getColVisible: () => true }
    }
    getRange(row: number, column: number, numRows = 1, numColumns = 1): unknown {
      return {
        getDisplayValues: () => [['wired']],
        getValues: () => [[undefined]],
        getCellStyleData: () => null,
        row,
        column,
        numRows,
        numColumns,
      }
    }
  }

  function wiredContext(composeStyle: (row: number, column: number) => unknown): PageLayoutContext {
    const active = new FakeSheet()
    const workbook = {
      getId: () => 'wb-1',
      getActiveSheet: () => active,
      getActiveRange: () => null,
      getSheets: () => [active],
    }
    const runtime = {
      univerAPI: { getActiveWorkbook: () => workbook },
      univer: {
        __getInjector: () => ({
          get: () => ({
            composeStyle: (_u: string, _s: string, row: number, col: number) =>
              composeStyle(row, col),
          }),
        }),
      },
    }
    const state = {
      file: { sheets: [], name: 'Book.xlsx', sessionId: 'session-1' },
      editJournal: createEditJournal(),
      sheetFilePageSetups: new Map(),
      flags: { preloadComplete: true, preloadRunning: false },
    } as unknown as LazyWorkbookState
    return {
      univerRef: { current: runtime as unknown as UniverRuntime },
      lazyWorkbookRef: { current: state },
      setMessage: () => {},
    } as unknown as PageLayoutContext
  }

  it('reads CF styles through the live service', async () => {
    const { request } = await buildPrintRequest(
      wiredContext(() => ({ style: { bg: { rgb: '#63be7b' } } })),
    )
    expect(request.html).toContain('background:#63be7b')
    expect(request.html).toContain('wired')
  })

  it('falls back to the plain worksheet without a live injector', async () => {
    const context = wiredContext(() => ({ style: { bg: { rgb: '#63be7b' } } }))
    ;(context.univerRef.current as unknown as { univer?: unknown }).univer = undefined
    const { request } = await buildPrintRequest(context)
    expect(request.html).toContain('wired')
    expect(request.html).not.toContain('#63be7b')
  })

  it('survives a composing service that throws per cell', async () => {
    const { request } = await buildPrintRequest(
      wiredContext(() => {
        throw new Error('boom')
      }),
    )
    expect(request.html).toContain('wired')
  })
})

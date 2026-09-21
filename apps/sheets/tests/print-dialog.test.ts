/**
 * The Print dialog's renderer plumbing: the shared print-request builder
 * applies the dialog's per-job overrides on top of the sheet's effective
 * page setup, and the File › Print / Ctrl+P wiring reaches the dialog from
 * the application menu, the preload allowlist, and the ribbon. The preview
 * channel's page-count scan and IPC shape live here too.
 */
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  BrowserWindow: class {},
  dialog: {},
}))
vi.mock('@airy-office/electron-utils', () => ({ showSaveDialogWithMemory: vi.fn() }))

import { countPdfPages } from '../src/main/pdf-export'
import { workbookPrintPreviewResultSchema } from '../src/shared/desktop-api'
import { controlsFromEffective } from '../src/renderer/PrintDialog'
import {
  buildActiveSheetPrintRequest,
  type PageLayoutContext,
} from '../src/renderer/page-layout-actions'
import { createEditJournal, recordPageSetup } from '../src/renderer/edit-journal'
import type { PrintWorksheet } from '../src/renderer/print-html'
import type { UniverRuntime } from '../src/renderer/univer-state'

function fakeWorksheet(): PrintWorksheet {
  return {
    getSheetName: () => 'Data',
    getLastRow: () => 2,
    getLastColumn: () => 1,
    getRowHeight: () => 20,
    getColumnWidth: () => 100,
    getSheet: () => ({ getRowVisible: () => true, getColVisible: () => true }),
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        [
          ['a', '1'],
          ['b', '2'],
        ]
          .slice(row, row + (numRows ?? 1))
          .map((cells) => cells.slice(column, column + (numColumns ?? 1))),
      getValues: () => [],
      getCellStyleData: () => null,
    })) as PrintWorksheet['getRange'],
  }
}

function layoutContext(): PageLayoutContext {
  const worksheet = fakeWorksheet()
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({
        getActiveSheet: () => ({
          ...worksheet,
          getSheetId: () => 'sheet-1',
          getSheetName: () => 'Data',
        }),
      }),
    },
  } as unknown as UniverRuntime
  return {
    univerRef: { current: runtime },
    lazyWorkbookRef: { current: null },
    setMessage: () => {},
  } as unknown as PageLayoutContext
}

const read = (relative: string): string => readFileSync(resolve(__dirname, relative), 'utf8')

describe('buildActiveSheetPrintRequest', () => {
  it('uses the sheet defaults when no overrides are given', async () => {
    const { request, effective } = await buildActiveSheetPrintRequest(layoutContext())
    expect(request.landscape).toBe(false)
    expect(request.pageSize).toBe('A4')
    expect(effective).toEqual({
      paperSize: 9,
      orientation: 'portrait',
      scale: 100,
      fitToPage: false,
      pageOrder: 'down-then-over',
    })
  })

  it('seeds from the journaled page order when the dialog passes none (BUG-1213)', async () => {
    const journal = createEditJournal()
    recordPageSetup(journal, 'sheet-1', { pageOrder: 'over-then-down' })
    const ctx = layoutContext()
    ctx.lazyWorkbookRef.current = {
      editJournal: journal,
      file: { sheets: [] },
      flags: { preloadComplete: true },
      sheetFilePageSetups: new Map(),
    } as unknown as PageLayoutContext['lazyWorkbookRef']['current']
    const { effective } = await buildActiveSheetPrintRequest(ctx)
    expect(effective.pageOrder).toBe('over-then-down')
  })

  it('applies the dialog overrides without touching the saved setup', async () => {
    const ctx = layoutContext()
    const { request, effective } = await buildActiveSheetPrintRequest(ctx, {
      paperSize: 1,
      orientation: 'landscape',
      scale: 50,
    })
    expect(request.landscape).toBe(true)
    expect(request.pageSize).toBe('Letter')
    expect(request.scale).toBe(0.5)
    expect(effective).toEqual({
      paperSize: 1,
      orientation: 'landscape',
      scale: 50,
      fitToPage: false,
      pageOrder: 'down-then-over',
    })
    // A second build without overrides still sees the untouched defaults.
    const again = await buildActiveSheetPrintRequest(ctx)
    expect(again.request.landscape).toBe(false)
    expect(again.effective.paperSize).toBe(9)
  })
})

describe('controlsFromEffective (dialog seed)', () => {
  it('seeds the controls from the sheet saved setup', () => {
    expect(
      controlsFromEffective({
        paperSize: 9,
        orientation: 'portrait',
        scale: 100,
        fitToPage: false,
        pageOrder: 'down-then-over',
      }),
    ).toEqual({
      paperSize: 9,
      orientation: 'portrait',
      scale: 100,
      fitToPage: false,
      pageOrder: 'down-then-over',
    })
    // A saved over-then-down seeds the dialog's order select.
    expect(
      controlsFromEffective({
        paperSize: 9,
        orientation: 'portrait',
        scale: 100,
        fitToPage: false,
        pageOrder: 'over-then-down',
      }),
    ).toEqual({
      paperSize: 9,
      orientation: 'portrait',
      scale: 100,
      fitToPage: false,
      pageOrder: 'over-then-down',
    })
    expect(
      controlsFromEffective({
        paperSize: 1,
        orientation: 'landscape',
        scale: 62.5,
        fitToPage: false,
        pageOrder: 'over-then-down',
      }),
    ).toEqual({
      paperSize: 1,
      orientation: 'landscape',
      scale: 63,
      fitToPage: false,
      pageOrder: 'over-then-down',
    })
  })

  it('fit-to-page owns the scale control at 100%', () => {
    expect(
      controlsFromEffective({
        paperSize: 9,
        orientation: 'portrait',
        scale: 40,
        fitToPage: true,
        pageOrder: 'down-then-over',
      }),
    ).toEqual({
      paperSize: 9,
      orientation: 'portrait',
      scale: 100,
      fitToPage: true,
      pageOrder: 'down-then-over',
    })
  })
})

describe('print preview page count', () => {
  it('counts /Type /Page objects and skips /Type /Pages tree nodes', () => {
    const pdf = Buffer.from(
      '<< /Type /Catalog /Pages 2 0 R >>\n' +
        '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>\n' +
        '<< /Type /Page /MediaBox [0 0 595 842] >>\n' +
        '<< /Type/Page/Parent 2 0 R >>\n',
      'latin1',
    )
    expect(countPdfPages(pdf)).toBe(2)
  })

  it('returns 0 when the scan finds nothing (the pdf-lib fallback path)', () => {
    expect(countPdfPages(Buffer.from('%PDF-1.4 stream-only bytes', 'latin1'))).toBe(0)
  })

  it('the preview IPC result carries only the page count', () => {
    expect(workbookPrintPreviewResultSchema.parse({ ok: true, pageCount: 3 })).toEqual({
      ok: true,
      pageCount: 3,
    })
    // the multi-MB base64 payload is gone from the channel — strict() rejects it
    expect(
      workbookPrintPreviewResultSchema.safeParse({
        ok: true,
        pageCount: 3,
        base64: 'AAAA',
      }).success,
    ).toBe(false)
    expect(workbookPrintPreviewResultSchema.parse({ ok: false, error: 'boom' })).toEqual({
      ok: false,
      error: 'boom',
    })
  })
})

describe('Print dialog wiring', () => {
  it('the application menu has Print with a Ctrl+P accelerator', () => {
    const mainSrc = read('../src/main/sheets-main.ts')
    expect(mainSrc).toMatch(
      /menuPrint[\s\S]{0,60}accelerator: 'CmdOrCtrl\+P'[\s\S]{0,40}sendMenuAction\('print'\)/,
    )
  })

  it('the preload allows the print menu action and the print channels', () => {
    const preloadSrc = read('../src/preload/index.ts')
    expect(preloadSrc).toContain("action === 'print'")
    expect(preloadSrc).toContain('IPC_CHANNELS.previewPrint')
    expect(preloadSrc).toContain('IPC_CHANNELS.print')
  })

  it('the menu action and the ribbon command open the print dialog', () => {
    const appSrc = read('../src/renderer/App.tsx')
    expect(appSrc).toMatch(/action === 'print'[\s\S]{0,60}setPrintDialogOpen\(true\)/)
    expect(appSrc).toMatch(/onOpenPrintDialog=\{\(\) => setPrintDialogOpen\(true\)\}/)
    const ribbonSrc = read('../src/renderer/ribbon-actions.ts')
    expect(ribbonSrc).toMatch(/command === 'print'[\s\S]{0,50}ctx\.openPrintDialog\(\)/)
  })

  it('the ribbon renders a File tab dropdown on non-mac platforms', () => {
    const shellSrc = read('../src/renderer/ExcelShell.tsx')
    expect(shellSrc).toMatch(/!IS_MAC &&[\s\S]{0,200}ribbon-tab-file/)
    expect(shellSrc).toContain("t('appFileOpen')")
    expect(shellSrc).toContain("t('appFileExportPdf')")
    expect(shellSrc).toContain("t('appFilePrint')")
  })

  it('Cmd/Ctrl+Y forwards to the redo menu action (Excel parity)', () => {
    const appSrc = read('../src/renderer/App.tsx')
    // capture-phase keydown on KeyY with the redo modifiers, sent as the
    // same menu action the Shift+Cmd+Z accelerator delivers
    expect(appSrc).toMatch(/event\.code === 'KeyY'[\s\S]{0,120}menuActionRef\.current\('redo'\)/)
    expect(appSrc).toMatch(/window\.addEventListener\('keydown', onRedoKey, true\)/)
    // and the shortcut sheet's listing matches the real wiring
    const registrySrc = read('../src/renderer/shortcut-registry.ts')
    expect(registrySrc).toContain("keys: '⇧⌘Z / ⌘Y'")
  })

  it('the File tab dropdown is dismissible and labeled as a menu button', () => {
    const shellSrc = read('../src/renderer/ExcelShell.tsx')
    // outside press / blur / chrome-press dismissal through the shared hook
    expect(shellSrc).toMatch(
      /useDismissablePopover\(fileMenuOpen, \(\) => setFileMenuOpen\(false\)/,
    )
    // Escape closes and returns focus to the toggle
    expect(shellSrc).toMatch(
      /fileMenuOpen[\s\S]{0,400}event\.key !== 'Escape'[\s\S]{0,300}fileTabButtonRef\.current\?\.focus\(\)/,
    )
    // menu-button semantics on the toggle
    expect(shellSrc).toMatch(/aria-haspopup="true"\s+aria-expanded=\{fileMenuOpen\}/)
  })
})

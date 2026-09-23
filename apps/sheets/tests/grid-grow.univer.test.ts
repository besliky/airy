/**
 * BUG-1615 regression: a new blank workbook is no longer trapped at its
 * 1000×26 default grid. Navigation (Name Box jumps, arrow keys at the edge)
 * grows the sheet on demand; growth stays out of the undo stack and the edit
 * journal (the save's <dimension> remains content-derived), and creating or
 * growing a book never mass-materializes cells.
 */
import {
  Direction,
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets/lib/facade'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '../src/gateway/csv-import'
import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  type CellEdit,
} from '../src/gateway/xlsx-gateway'
import {
  ensureSheetGrid,
  growOnArrowAtEdge,
  growSheetToRef,
  installEdgeNavigationGrowth,
  SHEET_MAX_COLUMNS,
  SHEET_MAX_ROWS,
} from '../src/renderer/grid-grow'
import { installJournalSuppressionUndoFilter } from '../src/renderer/univer-state'
import { createUniver } from '../src/renderer/create-univer'

function runtime(id: string, rowCount = 1000, columnCount = 26) {
  const rt = createUniver({
    locale: LocaleType.EN_US,
    locales: {},
    presets: [
      { plugins: [UniverFormulaEnginePlugin] },
      { plugins: [UniverSheetsPlugin] },
      { plugins: [UniverSheetsFormulaPlugin] },
    ],
  })
  rt.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id,
    name: 'wb',
    sheetOrder: ['main', 'other'],
    styles: {},
    sheets: {
      main: { id: 'main', name: 'Main', rowCount, columnCount, cellData: {} },
      other: { id: 'other', name: 'Other', rowCount, columnCount, cellData: {} },
    },
  })
  rt.univer.__getInjector().get(IUniverInstanceService).focusUnit(id)
  const workbook = rt.univerAPI.getActiveWorkbook()!
  const worksheet = workbook.getSheetBySheetId('main')!
  return { rt, workbook, worksheet }
}

function materializedCells(worksheet: {
  getSheet(): { getCellMatrix(): { getMatrix(): unknown } }
}): number {
  const matrix = (
    worksheet.getSheet().getCellMatrix() as {
      getMatrix(): Record<string, Record<string, unknown>>
    }
  ).getMatrix()
  return Object.values(matrix).reduce((sum, row) => sum + Object.keys(row).length, 0)
}

describe('ensureSheetGrid', () => {
  it('grows rows and columns past the 1000×26 default with a step', () => {
    const { rt, workbook, worksheet } = runtime('wb1')
    const before = { rows: worksheet.getMaxRows(), columns: worksheet.getMaxColumns() }
    const grown = ensureSheetGrid(rt, workbook, worksheet, 1500, 40)
    expect(before).toEqual({ rows: 1000, columns: 26 })
    expect(grown.rows).toBeGreaterThanOrEqual(1500)
    expect(grown.columns).toBeGreaterThanOrEqual(40)
    // idempotent within the grown extent
    const again = ensureSheetGrid(rt, workbook, worksheet, 1500, 40)
    expect(again).toEqual(grown)
  })

  it('treats 0 as no requirement and never shrinks the grid', () => {
    const { rt, workbook, worksheet } = runtime('wb2')
    const grown = ensureSheetGrid(rt, workbook, worksheet, 1500, 0)
    expect(grown.rows).toBeGreaterThanOrEqual(1500)
    expect(grown.columns).toBe(26)
    const kept = ensureSheetGrid(rt, workbook, worksheet, 10, 10)
    expect(kept).toEqual(grown)
  })

  it('clamps growth at the Excel sheet limits', () => {
    const { rt, workbook, worksheet } = runtime('wb3')
    const grown = ensureSheetGrid(
      rt,
      workbook,
      worksheet,
      SHEET_MAX_ROWS + 10,
      SHEET_MAX_COLUMNS + 5,
    )
    expect(grown.rows).toBe(SHEET_MAX_ROWS)
    expect(grown.columns).toBe(SHEET_MAX_COLUMNS)
  })

  it('grows without mass-materializing cells', () => {
    const { rt, workbook, worksheet } = runtime('wb4')
    ensureSheetGrid(rt, workbook, worksheet, 100_000, 100)
    expect(materializedCells(worksheet)).toBe(0)
    worksheet.getRange(99_999, 0).setValue('deep')
    expect(materializedCells(worksheet)).toBe(1)
  })
})

describe('growSheetToRef (Name Box growth)', () => {
  it('grows to fit a plain cell address and returns the target sheet', () => {
    const { rt, worksheet } = runtime('wb5')
    const target = growSheetToRef(rt, 'A1500')
    expect(target?.getSheetId()).toBe('main')
    expect(worksheet.getMaxRows()).toBeGreaterThanOrEqual(1500)
    // the jump now lands: the write the audit saw dropped silently succeeds
    worksheet.getRange(1499, 0).setValue('deep')
    expect(worksheet.getRange(1499, 0).getValue()).toBe('deep')
  })

  it('grows to fit ranges, column spans, and row spans', () => {
    const { rt, worksheet } = runtime('wb6')
    growSheetToRef(rt, 'Z1500:AB1600')
    expect(worksheet.getMaxRows()).toBeGreaterThanOrEqual(1600)
    expect(worksheet.getMaxColumns()).toBeGreaterThanOrEqual(28)
    growSheetToRef(rt, '1500:1700')
    expect(worksheet.getMaxRows()).toBeGreaterThanOrEqual(1700)
    growSheetToRef(rt, 'AB:AC')
    expect(worksheet.getMaxColumns()).toBeGreaterThanOrEqual(28)
  })

  it('grows the prefixed sheet, not the active one', () => {
    const { rt, workbook, worksheet } = runtime('wb7')
    const other = workbook.getSheetBySheetId('other')!
    const target = growSheetToRef(rt, 'Other!B1500')
    expect(target?.getSheetId()).toBe('other')
    expect(other.getMaxRows()).toBeGreaterThanOrEqual(1500)
    expect(worksheet.getMaxRows()).toBe(1000)
  })

  it('returns null for an unknown sheet prefix without growing anything', () => {
    const { rt, worksheet } = runtime('wb8')
    expect(growSheetToRef(rt, 'Nope!A1500')).toBeNull()
    expect(worksheet.getMaxRows()).toBe(1000)
  })
})

describe('arrow-key edge growth', () => {
  // The move-selection command itself lives in Univer's sheets-ui package
  // (not registered in the headless composition), so the tests drive the
  // before-command gate directly; moving within the grown grid is Univer's
  // stock behavior on any large sheet.
  it('grows the grid when a down arrow fires at the last row', () => {
    const { rt, workbook, worksheet } = runtime('wb9')
    workbook.setActiveRange(worksheet.getRange('A1000'))
    growOnArrowAtEdge(rt, 'sheet.command.move-selection', Direction.DOWN)
    expect(worksheet.getMaxRows()).toBeGreaterThanOrEqual(1001)
  })

  it('grows the grid when a right arrow fires at the last column', () => {
    const { rt, workbook, worksheet } = runtime('wb10')
    workbook.setActiveRange(worksheet.getRange('Z5'))
    growOnArrowAtEdge(rt, 'sheet.command.move-selection', Direction.RIGHT)
    expect(worksheet.getMaxColumns()).toBeGreaterThanOrEqual(27)
  })

  it('never grows upward, leftward, or for other commands', () => {
    const { rt, workbook, worksheet } = runtime('wb11')
    workbook.setActiveRange(worksheet.getRange('A1000'))
    growOnArrowAtEdge(rt, 'sheet.command.move-selection', Direction.UP)
    growOnArrowAtEdge(rt, 'sheet.command.move-selection', Direction.LEFT)
    growOnArrowAtEdge(rt, 'sheet.command.set-range-values', Direction.DOWN)
    expect(worksheet.getMaxRows()).toBe(1000)
    expect(worksheet.getMaxColumns()).toBe(26)
  })

  it('ignores mid-sheet down arrows', () => {
    const { rt, workbook, worksheet } = runtime('wb11b')
    workbook.setActiveRange(worksheet.getRange('A500'))
    growOnArrowAtEdge(rt, 'sheet.command.move-selection', Direction.DOWN)
    expect(worksheet.getMaxRows()).toBe(1000)
  })
})

describe('growth and undo', () => {
  it('keeps the grown grid when the typed value is undone', async () => {
    installJournalSuppressionUndoFilter()
    const { rt, worksheet } = runtime('wb12')
    growSheetToRef(rt, 'A1500')
    worksheet.getRange(1499, 0).setValue('typed')
    const grownRows = worksheet.getMaxRows()
    expect(grownRows).toBeGreaterThanOrEqual(1500)
    await rt.univerAPI.undo()
    expect(worksheet.getRange(1499, 0).getValue()).toBeNull()
    // the grid itself is model plumbing: undoing the typing must not shrink it
    expect(worksheet.getMaxRows()).toBe(grownRows)
  })

  it('the installed before-command gate tolerates unrelated commands', async () => {
    const { rt, workbook, worksheet } = runtime('wb13')
    const disposable = installEdgeNavigationGrowth(rt)
    const commandService = rt.univer.__getInjector().get(ICommandService)
    await commandService.executeCommand('sheet.command.set-range-values', {
      unitId: workbook.getId(),
      subUnitId: 'main',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      value: { 0: { 0: { v: 'kept' } } },
    })
    expect(worksheet.getRange(0, 0).getValue()).toBe('kept')
    expect(worksheet.getMaxRows()).toBe(1000)
    disposable.dispose()
  })
})

describe('save stays content-sized after growth (BUG-1615 DoD)', () => {
  it('writes a content-derived dimension and no empty row tails', async () => {
    const edit: CellEdit = {
      sheetName: 'Sheet1',
      row: 1499,
      column: 0,
      writeValue: true,
      cell: { value: 'deep' },
    }
    const mutation = await applyCellEditsToXlsx(await blankXlsxBuffer(), [edit])
    assertOnlyTouchedEntriesChanged(mutation)
    const zip = await JSZip.loadAsync(mutation.buffer)
    const worksheet = zip.file('xl/worksheets/sheet1.xml')
    expect(worksheet).toBeTruthy()
    const xml = (await worksheet!.async('text')) as string
    // openpyxl/Excel convention: dimension covers the used range only
    expect(xml).toContain('<dimension ref="A1:A1500"/>')
    const rows = [...xml.matchAll(/<row r="([0-9]+)"/g)].map((match) => Number(match[1]))
    expect(rows).toEqual([1500])
    expect(xml).not.toContain('<dimension ref="A1:XFD1048576"/>')
  })
})

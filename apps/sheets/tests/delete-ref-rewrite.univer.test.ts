/**
 * Headless end-to-end check of the delete→#REF! flow the BeforeCommandExecute
 * gate drives: cross-sheet dependent formulas are rewritten synchronously
 * BEFORE the remove-row command runs (so Univer carries the rewritten texts
 * verbatim through its dependent-cell relocation), the deletion succeeds
 * instead of being refused, and one undo — the batching the gate wraps around
 * rewrites, then the deletion — restoring texts, then rows and positions.
 */
import {
  ICommandService,
  IUndoRedoService,
  IUniverInstanceService,
  LocaleType,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets/lib/facade'
import { describe, expect, it } from 'vitest'

import { createUniver } from '../src/renderer/create-univer'
import {
  applyCrossSheetRewrites,
  collectCrossSheetDependentRewrites,
  deleteSpanSpec,
} from '../src/renderer/delete-ref-rewrite'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

function lazyState(): LazyWorkbookState {
  return {
    file: { sheets: [{ id: 'other', name: 'Other' }], visuals: [] },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { added: new Set(), removed: new Set() },
    },
    formulaText: new Map(),
  } as unknown as LazyWorkbookState
}

describe('cross-sheet #REF! rewrite around remove-row', () => {
  it('rewrites before the command and keeps one-step undo', async () => {
    const runtime = createUniver({
      locale: LocaleType.EN_US,
      locales: {},
      presets: [
        { plugins: [UniverFormulaEnginePlugin] },
        { plugins: [UniverSheetsPlugin] },
        { plugins: [UniverSheetsFormulaPlugin] },
      ],
    })
    runtime.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb1',
      sheetOrder: ['main', 'other'],
      name: 'wb',
      styles: {},
      sheets: {
        main: { id: 'main', name: 'Main', rowCount: 20, columnCount: 8, cellData: {} },
        other: { id: 'other', name: 'Other', rowCount: 20, columnCount: 8, cellData: {} },
      },
    })
    runtime.univer.__getInjector().get(IUniverInstanceService).focusUnit('wb1')
    const commandService = runtime.univer.__getInjector().get(ICommandService)
    const setValues = (subUnitId: string, value: Record<number, Record<number, unknown>>) =>
      commandService.executeCommand('sheet.command.set-range-values', {
        unitId: 'wb1',
        subUnitId,
        range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 7 },
        value,
      })
    await setValues('other', {
      1: { 1: { v: 10 }, 2: { v: 20 } },
      2: { 1: { v: 11 }, 2: { v: 21 } },
    })
    // A5: wholly-inside cross-sheet range; B5: partial; C5: same-sheet ref
    await setValues('main', {
      4: { 0: { f: '=SUM(Other!B2:B3)' }, 1: { f: '=SUM(Other!B2:B4)' }, 2: { f: '=B3+C3' } },
      3: { 1: { v: 5 }, 2: { v: 6 } },
    })
    const workbook = runtime.univerAPI.getActiveWorkbook()!
    const main = workbook.getSheetBySheetId('main')!
    const spec = deleteSpanSpec(lazyState(), () => 'Other', {
      op: 'delete_rows',
      sheetId: 'other',
      row: 2,
      count: 2,
    })

    // The gate's flow: run the deletion, then re-collect from the
    // post-deletion model (Univer relocated the dependents but left their
    // texts stale) and apply the rewrites.
    await commandService.executeCommand('sheet.command.remove-row', {
      unitId: 'wb1',
      subUnitId: 'other',
      range: { startRow: 1, endRow: 2, startColumn: 0, endColumn: 7 },
    })
    const rewrites = collectCrossSheetDependentRewrites(lazyState(), workbook, spec)
    expect(rewrites).toEqual([
      { sheetId: 'main', row: 2, column: 0, formula: '=SUM(#REF!)' },
      // B2:B4 minus deleted rows 2-3 leaves the old row-4 survivor at B2
      { sheetId: 'main', row: 2, column: 1, formula: '=SUM(Other!B2:B2)' },
    ])
    const rewriteBatch = runtime.univer.__getInjector().get(IUndoRedoService)
    const batch = rewriteBatch.__tempBatchingUndoRedo('wb1')
    try {
      applyCrossSheetRewrites(runtime, rewrites)
    } finally {
      batch.dispose()
    }

    // Univer relocated the dependent cells by the deletion delta (its
    // long-standing cross-sheet divergence) but carried the rewritten texts
    // verbatim — no stale or wrongly-shifted references remain.
    expect(main.getRange(2, 0).getFormulas()[0]?.[0]).toBe('=SUM(#REF!)')
    expect(main.getRange(2, 1).getFormulas()[0]?.[0]).toBe('=SUM(Other!B2:B2)')
    // Univer's own same-sheet rewriting: =B3+C3 referenced the deleted rows
    expect(main.getRange(2, 2).getFormulas()[0]?.[0]).toBe('=#REF!+#REF!')

    // ⌘Z reverts all the rewrites at once (they batch into one item): the
    // original texts return, still at the relocated rows. A second press
    // restores the deleted rows and the formulas' original positions.
    await runtime.univerAPI.undo()
    expect(main.getRange(2, 0).getFormulas()[0]?.[0]).toBe('=SUM(Other!B2:B3)')
    expect(main.getRange(2, 1).getFormulas()[0]?.[0]).toBe('=SUM(Other!B2:B4)')
    await runtime.univerAPI.undo()
    expect(main.getRange(4, 0).getFormulas()[0]?.[0]).toBe('=SUM(Other!B2:B3)')
    expect(main.getRange(4, 1).getFormulas()[0]?.[0]).toBe('=SUM(Other!B2:B4)')
    expect(main.getRange(4, 2).getFormulas()[0]?.[0]).toBe('=B3+C3')
    const other = workbook.getSheetBySheetId('other')!
    expect(other.getRange(1, 1).getValue()).toBe(10)
  })
})

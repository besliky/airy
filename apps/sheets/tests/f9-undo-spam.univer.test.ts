/**
 * UX-1763 at the model level: F9 spam (Calculate Now x50) between an edit and
 * its undo must not desynchronize the Univer undo stack — one Ctrl+Z has to
 * revert the edit, and the formula caches must roll back together with the
 * content (no dependent keeps a value computed from the undone state).
 *
 * Calculate Now dispatches SetTriggerFormulaCalculationStartMutation (a
 * mutation, not a command), so the recalculation itself must never land on
 * the undo stack — this pins that invariant: the stack holds exactly the
 * user's edit, undo pops it, and the recalculated caches of the dependents
 * re-converge to the restored values (Shift+F9's calculateSheet takes the
 * same direct-mutation path).
 *
 * Model-only boot (no UI plugin): the undo service and the formula engine
 * are unit-level machinery, the canvas is not involved. (The audit's
 * observation behind UX-1763 — one Ctrl+Z after F9-spam not reverting the
 * preceding edit — reproduces only in the docs app's tiptap history, see
 * .orchestrator/LOGS/UX-1762.md; sheets' Univer stack is clean.)
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import '@univerjs/engine-formula/lib/facade'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/lib/facade'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { describe, expect, it } from 'vitest'

import { calculateNow, calculateSheet } from '../src/renderer/calc-options'
import { createUniver } from '../src/renderer/create-univer'
import { undoStackDepth } from '../src/renderer/undo-carry'
import type { UniverRuntime } from '../src/renderer/univer-state'

function boot(): UniverRuntime {
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
    id: 'wb',
    name: 'wb',
    styles: {},
    sheets: {
      main: {
        id: 'main',
        name: 'Main',
        rowCount: 10,
        columnCount: 5,
        cellData: {
          // A1 = 10, B1 = A1*2 (cached 20), A2 = 20, A3 = SUM(A1:A2) (cached 30)
          0: { 0: { v: 10 }, 1: { f: '=A1*2', v: 20 } },
          1: { 0: { v: 20 } },
          2: { 0: { f: '=SUM(A1:A2)', v: 30 } },
        },
      },
    },
  })
  runtime.univer.__getInjector().get(IUniverInstanceService).focusUnit('wb')
  return runtime
}

describe('F9 spam keeps undo consistent (UX-1763)', () => {
  it('one undo reverts the edit after 50 recalculations', async () => {
    const rt = boot()
    const workbook = rt.univerAPI.getActiveWorkbook()!
    const main = workbook.getSheetBySheetId('main')!
    const valueAt = (row: number, column: number) => main.getRange(row, column).getValue()

    // The user's edit (SetRangeValuesCommand -> exactly one undo entry).
    await rt.univer
      .__getInjector()
      .get(ICommandService)
      .executeCommand('sheet.command.set-range-values', {
        unitId: 'wb',
        subUnitId: 'main',
        range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        value: { 0: { 0: { v: 100 } } },
      })
    expect(valueAt(0, 0)).toBe(100)

    // F9 spam x50, like the audit's driver, plus one Shift+F9 (the direct
    // set-formula-calculation-start dispatch of calculateSheet).
    for (let i = 0; i < 50; i += 1) calculateNow(rt)
    calculateSheet(rt)
    await rt.univerAPI.getFormula().onCalculationEnd()

    // The recalc is honest: caches reflect the edited value.
    expect(valueAt(2, 0)).toBe(120) // SUM(A1:A2) with A1 = 100
    expect(valueAt(0, 1)).toBe(200) // A1*2

    // The spam itself must not sit on the undo stack.
    expect(undoStackDepth(rt)).toBe(1)

    // One Ctrl+Z.
    await rt.univerAPI.undo()
    expect(valueAt(0, 0)).toBe(10)
    // The caches roll back with the content: no dependent keeps an F9 value
    // computed from the undone state.
    await rt.univerAPI.getFormula().onCalculationEnd()
    expect(valueAt(2, 0)).toBe(30)
    expect(valueAt(0, 1)).toBe(20)
  })
})

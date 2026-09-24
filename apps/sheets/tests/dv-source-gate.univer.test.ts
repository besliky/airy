/**
 * BUG-1601: a list validation whose source is a range on another (often
 * hidden) sheet — `=List!$C$1:$C$46` with stop semantics — must not block
 * input while the workbook is still streaming in: Univer resolves the rule
 * against an empty list until the source values materialize, so even valid
 * entries were rejected with the stop dialog. Excel ignores a rule whose
 * reference it cannot resolve. The source gate (data-validation-source-gate)
 * passes unresolved rules through and reactivates them once the source sheet
 * finishes loading; these headless checks drive the real Univer validation
 * stack.
 */
import {
  BooleanNumber,
  DataValidationStatus,
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  UniverInstanceType,
} from '@univerjs/core'
import { DataValidatorRegistryService, UniverDataValidationPlugin } from '@univerjs/data-validation'
import { FormulaResultStatus, UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import {
  DataValidationFormulaService,
  SheetsDataValidationValidatorService,
  SheetDataValidationModel,
  UniverSheetsDataValidationPlugin,
} from '@univerjs/sheets-data-validation'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets/lib/facade'
import { describe, expect, it, vi } from 'vitest'

import { refreshDvSourcesAfterSheetLoad } from '../src/renderer/data-validation-source-gate'
import { installDvSourcePassThrough } from '../src/renderer/data-validation-source-gate'
import { toUniverDvRule } from '../src/renderer/univer-sync'
import { createUniver } from '../src/renderer/create-univer'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

/// Unique per runtime: Univer keys the data-validation caches by unitId and
/// the cache services outlive individual Univer instances in one process.
let UNIT_ID = 'file-abc1'
let runtimeSeq = 0
const MAIN = 'main'
const LIST = 'list'

/// The streaming bookkeeping the gate reads: the Main window is loaded, the
/// hidden List sheet has not been touched yet.
function lazyState(): LazyWorkbookState {
  return {
    file: {
      sha256: 'abc123',
      sheets: [
        { id: MAIN, name: 'Main' },
        { id: LIST, name: 'List' },
      ],
    },
    generation: 1,
    loadedRanges: new Map([[MAIN, { startRow: 0, endRow: 99, startColumn: 0, endColumn: 9 }]]),
    flags: { preloadComplete: false, preloadRunning: false },
    editJournal: { structuralOps: new Map() },
  } as unknown as LazyWorkbookState
}

function createRuntime(state: LazyWorkbookState): UniverRuntime {
  UNIT_ID = `file-abc${++runtimeSeq}`
  state.file.sha256 = UNIT_ID.slice('file-'.length)
  const runtime = createUniver({
    locale: LocaleType.EN_US,
    locales: {},
    presets: [
      { plugins: [UniverFormulaEnginePlugin] },
      { plugins: [UniverSheetsPlugin] },
      { plugins: [UniverSheetsFormulaPlugin] },
      { plugins: [UniverDataValidationPlugin, UniverSheetsDataValidationPlugin] },
    ],
  })
  runtime.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: UNIT_ID,
    name: 'book',
    sheetOrder: [MAIN, LIST],
    styles: {},
    sheets: {
      [MAIN]: { id: MAIN, name: 'Main', rowCount: 100, columnCount: 10, cellData: {} },
      [LIST]: {
        id: LIST,
        name: 'List',
        rowCount: 50,
        columnCount: 4,
        hidden: BooleanNumber.TRUE,
        cellData: {},
      },
    },
  })
  runtime.univer.__getInjector().get(IUniverInstanceService).focusUnit(UNIT_ID)
  installDvSourcePassThrough(runtime, { current: state })
  return runtime
}

async function installListRule(runtime: UniverRuntime, formula: string): Promise<void> {
  const mapped = toUniverDvRule(
    {
      ranges: [{ startRow: 1, startColumn: 1, endRow: 40, endColumn: 1 }],
      ruleType: 'list',
      formulas: [formula],
      allowBlank: true,
      suppressDropdown: false,
      showInputMessage: false,
      showErrorMessage: true,
      errorStyle: 'stop',
    },
    'dv-rule-1',
  )
  await runtime.univer
    .__getInjector()
    .get(ICommandService)
    .executeCommand('data-validation.mutation.addRule', {
      unitId: UNIT_ID,
      subUnitId: MAIN,
      rule: mapped,
    })
}

async function setCell(
  runtime: UniverRuntime,
  subUnitId: string,
  row: number,
  col: number,
  value: string,
): Promise<void> {
  await runtime.univer
    .__getInjector()
    .get(ICommandService)
    .executeCommand('sheet.command.set-range-values', {
      unitId: UNIT_ID,
      subUnitId,
      range: { startRow: row, endRow: row, startColumn: col, endColumn: col },
      value: { [row]: { [col]: { v: value } } },
    })
}

async function validate(runtime: UniverRuntime, row: number): Promise<DataValidationStatus> {
  return runtime.univer
    .__getInjector()
    .get(SheetsDataValidationValidatorService)
    .validatorCell(UNIT_ID, MAIN, row, 1)
}

/// The engine registers the rule's formula in a debounced background cycle
/// (100 ms); poll the same mirror the gate itself consults — registered, with
/// every present formula result computed (no WAIT, no missing result) —
/// instead of sleeping a fixed window (PERF-1642: the old 600 ms settle per
/// test was pure wall). The mirror carries a slot per formula (formula2 stays
/// null here), so the check mirrors the gate's own vouchable predicate rather
/// than demanding both slots. Literal-list rules never reach the formula
/// engine, so they take no wait; if registration ever went missing, the
/// assertions below fail loudly instead of passing vacuously.
interface FormulaMirrorResult {
  status?: number
  result?: unknown
}

async function waitForRuleComputed(runtime: UniverRuntime): Promise<void> {
  const formulaService = runtime.univer.__getInjector().get(DataValidationFormulaService)
  await vi.waitFor(
    () => {
      const results = formulaService.getRuleFormulaResultSync(UNIT_ID, MAIN, 'dv-rule-1') as
        (FormulaMirrorResult | undefined)[] | undefined
      const computed =
        results != null &&
        !results.some(
          (item) =>
            item != null && (item.status === FormulaResultStatus.WAIT || item.result === undefined),
        )
      expect(computed, 'rule formula result not computed by the engine yet').toBe(true)
    },
    { timeout: 5_000, interval: 50 },
  )
}

/// The engine recomputes registered data-validation formulas in a debounced
/// background cycle; poll the resolved list (via the stock validator, exactly
/// what the dropdown reads) instead of sleeping a fixed amount.
async function waitForEngineList(runtime: UniverRuntime, expected: string[]): Promise<void> {
  const registry = runtime.univer.__getInjector().get(DataValidatorRegistryService)
  const listValidator = registry.getValidatorItem('list') as unknown as {
    getList(rule: unknown, unitId: string, subUnitId: string): string[]
  }
  const rule = runtime.univer
    .__getInjector()
    .get(SheetDataValidationModel)
    .getRules(UNIT_ID, MAIN)[0]
  await vi.waitFor(
    () => {
      expect(listValidator.getList(rule, UNIT_ID, MAIN)).toEqual(expected)
    },
    { timeout: 5_000, interval: 100 },
  )
}

describe('data-validation source gate (BUG-1601)', () => {
  it('passes a valid value while the hidden source sheet has not streamed in', async () => {
    const state = lazyState()
    const runtime = createRuntime(state)
    await installListRule(runtime, 'List!$C$1:$C$46')
    await waitForRuleComputed(runtime)
    await setCell(runtime, MAIN, 1, 1, 'CIF')
    // Stock validation resolves the reference to an empty list and would stop
    // the input; the gate reports VALID until the source is honest to check.
    await expect(validate(runtime, 1)).resolves.toBe(DataValidationStatus.VALID)
  }, 20_000)

  it('keeps literal-list rules enforcing during the streaming window', async () => {
    const state = lazyState()
    const runtime = createRuntime(state)
    await installListRule(runtime, '"CFR,CIF,CIP,CPT"')
    await setCell(runtime, MAIN, 1, 1, 'BOGUS')
    await expect(validate(runtime, 1)).resolves.toBe(DataValidationStatus.INVALID)
  }, 20_000)

  it('reactivates against the resolved source: valid passes, invalid stops', async () => {
    const state = lazyState()
    const runtime = createRuntime(state)
    await installListRule(runtime, 'List!$C$1:$C$46')
    await waitForRuleComputed(runtime)
    // The preload block for the hidden List sheet completes: its values land
    // and the coverage ledger declares the sheet materialized.
    await setCell(runtime, LIST, 0, 2, 'CFR')
    await setCell(runtime, LIST, 1, 2, 'CIF')
    await setCell(runtime, LIST, 2, 2, 'CIP')
    await setCell(runtime, LIST, 3, 2, 'CPT')
    state.loadedRanges.set(LIST, { startRow: 0, endRow: 49, startColumn: 0, endColumn: 3 })
    refreshDvSourcesAfterSheetLoad(runtime, state, LIST)
    await waitForEngineList(runtime, ['CFR', 'CIF', 'CIP', 'CPT'])
    await setCell(runtime, MAIN, 1, 1, 'CIF')
    await expect(validate(runtime, 1)).resolves.toBe(DataValidationStatus.VALID)
    await setCell(runtime, MAIN, 2, 1, 'BOGUS')
    await expect(validate(runtime, 2)).resolves.toBe(DataValidationStatus.INVALID)
  }, 20_000)

  it('resolves the hidden-sheet source for the dropdown list after full load', async () => {
    const state = lazyState()
    const runtime = createRuntime(state)
    await installListRule(runtime, 'List!$C$1:$C$46')
    await waitForRuleComputed(runtime)
    const registry = runtime.univer.__getInjector().get(DataValidatorRegistryService)
    const listValidator = registry.getValidatorItem('list') as unknown as {
      getList(rule: unknown, unitId: string, subUnitId: string): string[]
    }
    const rule = runtime.univer
      .__getInjector()
      .get(SheetDataValidationModel)
      .getRules(UNIT_ID, MAIN)[0]
    // While the source is unloaded the dropdown list reads empty.
    expect(listValidator.getList(rule, UNIT_ID, MAIN)).toEqual([])
    await setCell(runtime, LIST, 0, 2, 'CFR')
    await setCell(runtime, LIST, 1, 2, 'CIF')
    await setCell(runtime, LIST, 2, 2, 'CIP')
    await setCell(runtime, LIST, 3, 2, 'CPT')
    state.loadedRanges.set(LIST, { startRow: 0, endRow: 49, startColumn: 0, endColumn: 3 })
    refreshDvSourcesAfterSheetLoad(runtime, state, LIST)
    await waitForEngineList(runtime, ['CFR', 'CIF', 'CIP', 'CPT'])
    expect(listValidator.getList(rule, UNIT_ID, MAIN)).toEqual(['CFR', 'CIF', 'CIP', 'CPT'])
  }, 20_000)

  it('enforces same-sheet range rules that are fully loaded, as before', async () => {
    const state = lazyState()
    const runtime = createRuntime(state)
    await setCell(runtime, MAIN, 0, 6, 'AA')
    await setCell(runtime, MAIN, 1, 6, 'BB')
    await installListRule(runtime, '$G$1:$G$4')
    await waitForEngineList(runtime, ['AA', 'BB'])
    await setCell(runtime, MAIN, 1, 1, 'AA')
    await expect(validate(runtime, 1)).resolves.toBe(DataValidationStatus.VALID)
    await setCell(runtime, MAIN, 2, 1, 'CC')
    await expect(validate(runtime, 2)).resolves.toBe(DataValidationStatus.INVALID)
  }, 20_000)
})

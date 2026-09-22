/**
 * Excel ignores a data-validation rule whose source reference it cannot
 * resolve (range on a closed workbook, #REF!): typing anything is accepted.
 * Univer instead evaluates such rules against an empty result while a
 * streamed workbook is still loading — a list source on another sheet,
 * commonly a hidden one (`=List!$C$1:$C$46`), is not materialized yet, every
 * value "fails", and a stop-style rule blocks the edit with a dialog
 * (BUG-1601: ~20 s of blocked input on large files). This module mirrors
 * Excel's semantics:
 *
 * - Pass-through: `SheetDataValidationModel.validator` reports VALID for a
 *   formula-driven rule whose source the model cannot honestly evaluate —
 *   referenced values not streamed in yet, or the engine result not
 *   computed. Literal-list rules and fully loaded sources validate as
 *   before.
 * - Reactivation: when a sheet's values finish loading (preload block done),
 *   `refreshDvSourcesAfterSheetLoad` marks the rules reading that sheet
 *   dirty in Univer's formula engine; the recalculated result flows through
 *   the stock `formulaResult$` → cache-invalidation → revalidation chain,
 *   and the rule enforces again (stop on invalid works).
 */
import {
  DataValidationStatus,
  IUniverInstanceService,
  UniverInstanceType,
  isFormulaString,
} from '@univerjs/core'
import {
  FormulaResultStatus,
  LexerTreeBuilder,
  RegisterOtherFormulaService,
  deserializeRangeWithSheet,
  isReferenceString,
  sequenceNodeType,
} from '@univerjs/engine-formula'
import {
  DataValidationFormulaService,
  SheetDataValidationModel,
} from '@univerjs/sheets-data-validation'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

interface DvRuleLike {
  uid: string
  formula1?: string
  formula2?: string
}

interface DvPos {
  unitId: string
  subUnitId: string
  row: number
  col: number
}

type DvOnCompete = (status: DataValidationStatus, changed: boolean) => void

interface DvCacheMatrix {
  realDeleteValue(row: number, col: number): void
}

interface DvValidatorThis {
  _dataValidationCacheService?: {
    ensureCache(unitId: string, subUnitId: string): DvCacheMatrix
  }
}

interface DirtyCellResult {
  status?: number
  result?: unknown
}

interface RegisteredFormulaInfo {
  id: string
}

/// Structural slice of Univer's Workbook the reference checks need.
interface WorkbookLike {
  getSheetBySheetId(sheetId: string): { getName(): string } | null | undefined
  getSheetBySheetName(name: string): { getSheetId(): string; getName(): string } | null | undefined
}

interface SourceTracker {
  runtime: UniverRuntime
  lazyWorkbookRef: { current: LazyWorkbookState | null }
}

let current: SourceTracker | null = null
let patchInstalled = false

/// The controller batches commands on a 100 ms debounce. Marking the formulas
/// dirty inside that same window makes the mark coalesce with the sheet's own
/// value-write cycle, and the merged cycle evaluates the formula against the
/// pre-write snapshot; issuing the mark after the debounce lands it in its own
/// cycle, which recomputes against the resolved values.
const REFRESH_MARK_DELAY_MS = 250
const REFRESH_MARK_RETRY_MS = 900

/// The patch must live on the class prototype: the injector hands out a lazy
/// redi proxy, so assigning a wrapper onto the resolved instance only shadows
/// the proxy (same shape as the undo filter in univer-state.ts).
export function installDvSourcePassThrough(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
): void {
  current = { runtime, lazyWorkbookRef }
  if (patchInstalled) return
  patchInstalled = true
  const proto = SheetDataValidationModel.prototype as unknown as {
    validator: (
      this: DvValidatorThis,
      rule: DvRuleLike,
      pos: DvPos,
      onCompete?: DvOnCompete,
    ) => DataValidationStatus
  }
  const original = proto.validator
  proto.validator = function (rule, pos, onCompete) {
    if (
      rule &&
      pos &&
      current !== null &&
      dvSourceUnresolved(current, pos.unitId, pos.subUnitId, rule)
    ) {
      // The streaming window may already have cached INVALID for this cell;
      // drop it so paint and cache reads reflect the pass-through.
      try {
        this._dataValidationCacheService
          ?.ensureCache(pos.unitId, pos.subUnitId)
          .realDeleteValue(pos.row, pos.col)
      } catch {
        // Cache shape drift must never break validation itself.
      }
      onCompete?.(DataValidationStatus.VALID, false)
      return DataValidationStatus.VALID
    }
    return original.call(this, rule, pos, onCompete)
  }
}

/// True when the rule reads a source the model cannot honestly evaluate yet:
/// literal-list rules (no `=` formulas) always enforce.
function dvSourceUnresolved(
  tracker: SourceTracker,
  unitId: string,
  subUnitId: string,
  rule: DvRuleLike,
): boolean {
  const state = tracker.lazyWorkbookRef.current
  // Only streamed file loads have partial sources; snapshot workbooks are
  // complete the moment they exist.
  if (!state || unitId !== `file-${state.file.sha256}`) return false
  if (state.flags.preloadComplete) return false
  const formulas = [rule.formula1, rule.formula2].filter(
    (formula): formula is string => typeof formula === 'string' && isFormulaString(formula),
  )
  if (formulas.length === 0) return false
  const injector = tracker.runtime.univer.__getInjector()
  const instances = injector.get(IUniverInstanceService)
  const workbook = instances.getUnit(unitId, UniverInstanceType.UNIVER_SHEET) as unknown as
    WorkbookLike | undefined
  if (!workbook) return false
  const lexer = injector.get(LexerTreeBuilder)
  if (
    formulas.some((formula) =>
      formulaSourceUnmaterialized(state, workbook, subUnitId, formula, lexer),
    )
  ) {
    return true
  }
  // The engine registered the rule's formula but has not computed it yet —
  // stock validation would read an empty result and reject everything.
  const results = injector
    .get(DataValidationFormulaService)
    .getRuleFormulaResultSync(unitId, subUnitId, rule.uid) as
    (DirtyCellResult | undefined)[] | undefined
  if (results === undefined) return false
  return results.some(
    (item) =>
      item != null && (item.status === FormulaResultStatus.WAIT || item.result === undefined),
  )
}

/// True when the formula's tokens cannot be vouched for against the loaded
/// model: references outside the streamed window, or INDIRECT/OFFSET/defined
/// names whose resolution the model state cannot confirm.
function formulaSourceUnmaterialized(
  state: LazyWorkbookState,
  workbook: WorkbookLike,
  ruleSubUnitId: string,
  formula: string,
  lexer: LexerTreeBuilder,
): boolean {
  const nodes = lexer.sequenceNodesBuilder(formula) ?? []
  for (const node of nodes) {
    if (typeof node === 'string') continue
    if (node.nodeType === sequenceNodeType.REFERENCE && isReferenceString(node.token)) {
      const reference = deserializeRangeWithSheet(node.token)
      if (
        referenceSourceUnmaterialized(
          state,
          workbook,
          ruleSubUnitId,
          reference.sheetName || null,
          reference.range,
        )
      ) {
        return true
      }
    } else if (
      node.nodeType === sequenceNodeType.FUNCTION ||
      node.nodeType === sequenceNodeType.DEFINED_NAME ||
      node.nodeType === sequenceNodeType.TABLE
    ) {
      return true
    }
  }
  return false
}

interface ReferenceRange {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

function referenceSourceUnmaterialized(
  state: LazyWorkbookState,
  workbook: WorkbookLike,
  ruleSubUnitId: string,
  sheetName: string | null,
  range: ReferenceRange,
): boolean {
  let sheetId = ruleSubUnitId
  if (sheetName !== null) {
    const fileSheet = state.file.sheets.find((sheet) => sheet.name === sheetName)
    if (fileSheet) {
      sheetId = fileSheet.id
    } else {
      const worksheet = workbook.getSheetBySheetName(sheetName)
      // A reference to a sheet that does not exist at all is broken — Excel
      // never blocks on it.
      if (!worksheet) return false
      sheetId = worksheet.getSheetId()
    }
  }
  // Structural row/column edits shift screen coordinates; the coverage ledger
  // and the file-space reference only coincide before the first shift, so a
  // shifted sheet stays skipped until the full load declares everything.
  if ((state.editJournal.structuralOps.get(sheetId)?.length ?? 0) > 0) return true
  const loaded = state.loadedRanges.get(sheetId)
  if (!loaded) return true
  return (
    range.startRow < loaded.startRow ||
    range.endRow > loaded.endRow ||
    range.startColumn < loaded.startColumn ||
    range.endColumn > loaded.endColumn
  )
}

/// After `sheetId`'s values fully materialized (preload block done), re-mark
/// every installed rule that reads this sheet so the engine recalculates their
/// source formulas from the now-resolved values. The stock `formulaResult$`
/// handler then invalidates the list/status caches and revalidates the cells,
/// and stop semantics apply again. Until the recalculation lands, the gate
/// above treats the re-marked (WAIT) formulas as unresolved and keeps passing
/// inputs through.
export function refreshDvSourcesAfterSheetLoad(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  sheetId: string,
): void {
  const unitId = `file-${state.file.sha256}`
  if (state.flags.preloadComplete) return
  const injector = runtime.univer.__getInjector()
  const instances = injector.get(IUniverInstanceService)
  const workbook = instances.getUnit(unitId, UniverInstanceType.UNIVER_SHEET) as unknown as
    WorkbookLike | undefined
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  if (!workbook || !worksheet) return
  const sheetName = worksheet.getName()
  const lexer = injector.get(LexerTreeBuilder)
  const model = injector.get(SheetDataValidationModel)
  const formulaService = injector.get(DataValidationFormulaService)
  const registerService = injector.get(RegisterOtherFormulaService) as {
    markFormulaDirty(unitId: string, subUnitId: string, formulaId: string): void
  }
  const formulaIds: { ruleSheetId: string; formulaId: string }[] = []
  for (const [ruleSheetId, rules] of model.getUnitRules(unitId) as [string, DvRuleLike[]][]) {
    for (const rule of rules) {
      if (!ruleReadsSheet(state, workbook, lexer, ruleSheetId, rule, sheetName, sheetId)) continue
      for (const info of (formulaService.getRuleFormulaInfo(unitId, ruleSheetId, rule.uid) ??
        []) as (RegisteredFormulaInfo | undefined)[]) {
        if (info?.id) formulaIds.push({ ruleSheetId, formulaId: info.id })
      }
    }
  }
  if (formulaIds.length === 0) return
  const markAll = (): void => {
    for (const { ruleSheetId, formulaId } of formulaIds) {
      registerService.markFormulaDirty(unitId, ruleSheetId, formulaId)
    }
  }
  // The formulas keep their first-calculation dependency entries only while
  // they stay registered — re-registering (remove + add) would drop them, so
  // the stock mark path is the only recompute trigger. The first mark lands
  // after the sheet's own value-write cycle; the retry covers a cycle that was
  // still running when the first mark fired.
  setTimeout(markAll, REFRESH_MARK_DELAY_MS)
  setTimeout(markAll, REFRESH_MARK_RETRY_MS)
}

/// True when any static reference of the rule targets `sheetId` (or its own
/// sheet), or when the formula carries opaque tokens (INDIRECT/defined names)
/// that could resolve there.
function ruleReadsSheet(
  state: LazyWorkbookState,
  workbook: WorkbookLike,
  lexer: LexerTreeBuilder,
  ruleSheetId: string,
  rule: DvRuleLike,
  sheetName: string,
  sheetId: string,
): boolean {
  const formulas = [rule.formula1, rule.formula2].filter(
    (formula): formula is string => typeof formula === 'string' && isFormulaString(formula),
  )
  if (formulas.length === 0) return false
  return formulas.some((formula) =>
    formulaReadsSheet(state, workbook, lexer, ruleSheetId, formula, sheetName, sheetId),
  )
}

function formulaReadsSheet(
  state: LazyWorkbookState,
  workbook: WorkbookLike,
  lexer: LexerTreeBuilder,
  ruleSheetId: string,
  formula: string,
  sheetName: string,
  sheetId: string,
): boolean {
  const nodes = lexer.sequenceNodesBuilder(formula) ?? []
  for (const node of nodes) {
    if (typeof node === 'string') continue
    if (node.nodeType === sequenceNodeType.REFERENCE && isReferenceString(node.token)) {
      const reference = deserializeRangeWithSheet(node.token)
      const target = reference.sheetName
        ? (state.file.sheets.find((sheet) => sheet.name === reference.sheetName)?.id ??
          workbook.getSheetBySheetName(reference.sheetName)?.getSheetId())
        : ruleSheetId
      if (target === sheetId) return true
    } else if (
      node.nodeType === sequenceNodeType.FUNCTION ||
      node.nodeType === sequenceNodeType.DEFINED_NAME ||
      node.nodeType === sequenceNodeType.TABLE
    ) {
      return true
    }
  }
  return false
}

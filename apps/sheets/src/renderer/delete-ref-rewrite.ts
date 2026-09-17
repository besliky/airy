/**
 * Excel-compatible #REF! handling for row/column deletion.
 *
 * Univer's live model already rewrites same-sheet references when a
 * remove-row/remove-col command executes (wholly-inside references become
 * #REF!, partially overlapping ranges clip). Two gaps remain:
 *  - formulas on OTHER sheets that reference the deleted sheet: Univer
 *    leaves their texts stale or shifts them into wrong ranges, and
 *    relocates the dependent cells by the deletion's row delta (a
 *    long-standing divergence from Excel that predates this module — the
 *    post-save reopen rebuilds the model from the corrected file);
 *  - the save's own structural replay used to abort on any wholly-inside
 *    reference (shiftFormulaText's fail-closed mode).
 *
 * The save side now emits #REF! (xlsx-structure's 'ref-error' mode). This
 * module fixes the model side: the BeforeCommandExecute gate collects the
 * affected cross-sheet formulas from the pristine pre-deletion model and
 * rewrites them SYNCHRONOUSLY BEFORE the command runs — Univer then carries
 * the rewritten texts verbatim through its relocation, because they no
 * longer reference the deleted span. The rewrites are journaled set-values
 * commands wrapped in one undo batch with the deletion, so a single ⌘Z
 * restores rows and formulas together; if the command never lands (canceled
 * by a later gate), the batch's only item is rolled back with one undo. The
 * harvested file-coordinate formula index (streamed workbooks' formula bar)
 * is rewritten in place once the deletion is confirmed; those keys live on
 * other sheets, which the deletion never moves.
 */
import { ICommandService } from '@univerjs/core'

import { columnIndex } from '../domain/cell-address'
import { shiftFormulaText, type Axis, type Shift } from '../gateway/xlsx-structure'
import { isSheetRemoved } from './edit-journal'
import type { LazyWorkbookState } from './univer-state'

export type DeleteSpanOp =
  | { op: 'delete_rows'; sheetId: string; row: number; count: number }
  | { op: 'delete_cols'; sheetId: string; column: string; count: number }

export interface DeletedSpanSpec {
  axis: Axis
  sheetId: string
  shift: Shift
  deletedSheetName: string
}

/// The deleted span in the shape the shared token rewriter consumes.
export function deleteSpanSpec(
  state: LazyWorkbookState,
  sheetNameOf: (sheetId: string) => string | undefined,
  operation: DeleteSpanOp,
): DeletedSpanSpec {
  const axis: Axis = operation.op === 'delete_cols' ? 'column' : 'row'
  const index = operation.op === 'delete_cols' ? columnIndex(operation.column) : operation.row - 1
  return {
    axis,
    sheetId: operation.sheetId,
    shift: {
      boundary: index,
      delta: -operation.count,
      deleted: { start: index, end: index + operation.count - 1 },
    },
    deletedSheetName:
      state.file.sheets.find((sheet) => sheet.id === operation.sheetId)?.name ??
      sheetNameOf(operation.sheetId) ??
      '',
  }
}

/// Excel rewrite of one formula text against the deleted span: wholly-inside
/// references become a bare #REF! token (the qualifier is dropped: the
/// recalc engine rejects `Sheet!#REF!` while Univer's own rewrite emits
/// the bare form too), partially
/// overlapping ranges clip, references beyond the span shift. Null when the
/// formula does not change.
export function rewriteFormulaForDeletedSpan(
  formula: string,
  spec: DeletedSpanSpec,
  qualifiedOnly: boolean,
): string | null {
  const rewritten = shiftFormulaText(
    formula,
    spec.deletedSheetName,
    spec.shift,
    spec.axis,
    qualifiedOnly,
    'ref-error',
  )
  return rewritten === formula ? null : rewritten
}

/// One model cell whose formula must become a rewritten text.
export interface CrossSheetRewrite {
  sheetId: string
  row: number
  column: number
  formula: string
}

/// Minimal workbook surface the scan needs (the Univer facade satisfies it).
export interface RewriteWorkbook {
  getSheets(): {
    getSheetId(): string
    getSheetName(): string
    getMaxRows(): number
    getMaxColumns(): number
    getRange(
      row: number,
      column: number,
      rows: number,
      columns: number,
    ): { getFormulas(): string[][] }
  }[]
}

/// Whether the deletion the rewrites were computed for actually landed in the
/// model: the sheet's row/column count shrank by the deleted span's size.
/// Univer emits CommandExecuted even when the command handler DECLINES
/// (protected sheet, invalid range), and the safety-valve timeout reaches the
/// same finish() path after a throw or cancel — without this check a declined
/// deletion would rewrite qualified references on other sheets to #REF!/clipped
/// forms via journaled commands and reach the next save.
export function deletionLanded(
  workbook: RewriteWorkbook,
  spec: DeletedSpanSpec,
  beforeCount: number,
  deletedCount: number,
): boolean {
  const sheet = workbook.getSheets().find((candidate) => candidate.getSheetId() === spec.sheetId)
  if (!sheet) return false
  const afterCount = spec.axis === 'row' ? sheet.getMaxRows() : sheet.getMaxColumns()
  return deletedCount > 0 && afterCount === beforeCount - deletedCount
}

/// Formulas on every OTHER sheet that reference the deleted span (only
/// sheet-qualified tokens can) and need a #REF!/clip rewrite in the model.
/// Same-sheet formulas are left to Univer's own remove-row/col rewriting.
export function collectCrossSheetDependentRewrites(
  state: LazyWorkbookState,
  workbook: RewriteWorkbook,
  spec: DeletedSpanSpec,
): CrossSheetRewrite[] {
  const rewrites: CrossSheetRewrite[] = []
  for (const sheet of workbook.getSheets()) {
    const sheetId = sheet.getSheetId()
    if (sheetId === spec.sheetId) continue
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const formulas = sheet.getRange(0, 0, sheet.getMaxRows(), sheet.getMaxColumns()).getFormulas()
    for (let row = 0; row < formulas.length; row += 1) {
      const columns = formulas[row]
      if (!columns) continue
      for (let column = 0; column < columns.length; column += 1) {
        const formula = columns[column]
        if (!formula) continue
        const rewritten = rewriteFormulaForDeletedSpan(formula, spec, true)
        if (rewritten !== null) {
          rewrites.push({ sheetId, row, column, formula: rewritten })
        }
      }
    }
  }
  return rewrites
}

/// Rewrites the harvested file-coordinate formula index for sheets other
/// than the deleted one: display-only (the keys never move — the deletion is
/// on another sheet), and the save rewrites the file's own formulas anyway.
/// A journal entry at the same cell supersedes the harvested text, so those
/// keys are left alone.
export function rewriteHarvestedFormulaTexts(
  state: LazyWorkbookState,
  spec: DeletedSpanSpec,
): void {
  for (const [sheetId, cells] of state.formulaText) {
    if (sheetId === spec.sheetId) continue
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const journalCells = state.editJournal.cells.get(sheetId)
    for (const [key, text] of cells) {
      if (journalCells?.has(key)) continue
      const rewritten = rewriteFormulaForDeletedSpan(text, spec, true)
      if (rewritten !== null) cells.set(key, rewritten)
    }
  }
}

/// Minimal runtime surface the apply step needs.
export interface RewriteRuntime {
  univer: { __getInjector(): unknown }
  univerAPI: { getActiveWorkbook(): { getId(): string } | null | undefined }
}

/// Writes each rewrite into the live model through the same set-range-values
/// command the cell editor uses — journaled, undoable, and synchronous so
/// the whole batch lands as undo items in a deterministic order.
export function applyCrossSheetRewrites(
  runtime: RewriteRuntime,
  rewrites: readonly CrossSheetRewrite[],
): void {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const commandService = (runtime.univer.__getInjector() as { get(token: unknown): unknown }).get(
    ICommandService,
  ) as { syncExecuteCommand(id: string, params?: unknown): unknown }
  for (const rewrite of rewrites) {
    commandService.syncExecuteCommand('sheet.command.set-range-values', {
      unitId,
      subUnitId: rewrite.sheetId,
      range: {
        startRow: rewrite.row,
        endRow: rewrite.row,
        startColumn: rewrite.column,
        endColumn: rewrite.column,
      },
      value: { [rewrite.row]: { [rewrite.column]: { f: rewrite.formula } } },
    })
  }
}

/// Inputs of the finish step the BeforeCommandExecute gate runs once the
/// deletion settles (CommandExecuted, or the safety valve on cancel/throw).
export interface FinishRewritesArgs {
  runtime: RewriteRuntime
  state: LazyWorkbookState
  workbook: RewriteWorkbook
  spec: DeletedSpanSpec
  /** sheet row/column count captured before the command ran */
  countBefore: number | undefined
  /** size of the requested deleted span */
  spanCount: number
  /** opens one undo item for all the rewrites; null/omitted while an AI bulk
   *  edit owns its own batch (injected to keep this module import-light) */
  beginBatch?: (() => { settle(): void }) | null
}

/// The finish step of the cross-sheet #REF! flow: verify the deletion
/// actually landed, then re-collect rewrites from the (relocated) model and
/// apply them as journaled commands under one undo item. Returns false —
/// rewrites dropped, a console.debug note left — when the deletion never
/// executed: Univer emits CommandExecuted even when the handler declines
/// (protected sheet, invalid range) and the safety valve reaches this step
/// after a throw or cancel, and rewriting formulas for a deletion that did
/// not happen would corrupt the workbook at the next save.
export function finishCrossSheetRewrites(args: FinishRewritesArgs): boolean {
  const { runtime, state, workbook, spec, countBefore, spanCount, beginBatch } = args
  if (countBefore === undefined || !deletionLanded(workbook, spec, countBefore, spanCount)) {
    console.debug('cross-sheet #REF! rewrite skipped: the deletion did not land')
    return false
  }
  const batch = beginBatch?.() ?? null
  try {
    applyCrossSheetRewrites(runtime, collectCrossSheetDependentRewrites(state, workbook, spec))
    rewriteHarvestedFormulaTexts(state, spec)
  } catch {
    // Best-effort model polish; the save's own #REF! emission remains the
    // backstop for file correctness.
  } finally {
    batch?.settle()
  }
  return true
}

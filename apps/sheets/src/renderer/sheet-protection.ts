/**
 * PAR-204: worksheet-protection emulation. Excel's model, mirrored here:
 *
 *  - The file's <sheetProtection> element carries a protection flag, an
 *    optional legacy password hash, and per-action attributes in raw OOXML
 *    polarity (true = the action is PREVENTED while protected; absent =
 *    schema default — the two select* actions default to allowed, every
 *    other one to prevented).
 *  - While a sheet is protected, locked cells refuse edits ("The cell or
 *    chart that you are trying to change is on a protected sheet…") and the
 *    modeled actions (insert/delete rows/columns, sort, autofilter, …) are
 *    refused unless their attribute says allowed.
 *  - Unprotecting verifies the typed password against the stored legacy hash;
 *    modern hashValue sheets fail closed (the agile hash is not verifiable
 *    here).
 *
 * Locked resolution order (Excel default = locked):
 *   1. the cell's journal style patch (a session "Unlock Cell" toggle),
 *   2. the cell's own file xf, delivered as the `custom.unlocked` install
 *      flag (see patchWorksheetRangeInner) and read back through the raw
 *      cell matrix — composed getCell results can drop the custom bag
 *      (BUG-1715),
 *   3. the column-default xf (<col style=>) from the file metadata,
 *   4. locked.
 * Row-default xfs are the documented gap: row styles stream without a
 * per-sheet record to consult at gate time, so a row-level unlocked default
 * still edits as locked (noted in .orchestrator/LOGS/PAR-204.md — a rare
 * authoring path; the Lock Cell ribbon toggle writes cell-level xfs).
 *
 */
import type { LazyWorkbookState } from './univer-state'
import type { SheetProtectionAttributes, WorkbookSheetProtection } from '../shared/desktop-api'
import { legacyPasswordMatches } from '../shared/legacy-password'
import {
  MOVE_COLS_COMMAND,
  MOVE_RANGE_COMMAND,
  MOVE_ROWS_COMMAND,
  OPEN_FILTER_PANEL_OPERATION,
  SET_RANGE_VALUES_COMMAND,
  SET_RANGE_VALUES_MUTATION,
  SORT_COMMAND_PATTERN,
  STRUCTURAL_EDIT_COMMAND_PATTERN,
  FILTER_COMMAND_PATTERN,
} from './app-constants'
import { screenToFile } from './view-transform'

/// The cell install flag baked into ICellData.custom when the file's own xf
/// says locked="0" (read back with worksheet.getCell(...).custom).
export const UNLOCKED_CELL_KEY = 'unlocked'

export type EffectiveSheetProtection = WorkbookSheetProtection

/// The journal delta wins over the file state; null when the sheet is unknown
/// (treated as unprotected — brand-new sheets carry no file protection).
export function effectiveSheetProtection(
  state: LazyWorkbookState,
  sheetId: string,
): EffectiveSheetProtection | null {
  const delta = state.editJournal.sheetProtection.get(sheetId)
  const file = state.sheetProtections.get(sheetId)
  if (!delta) return file ?? null
  const base = file ?? { protected: false, hasPassword: false }
  return {
    ...base,
    ...delta.attributes,
    protected: delta.protected,
    hasPassword: delta.passwordHash != null || (delta.protected && base.hasPassword),
    ...(delta.passwordHash != null ? { passwordHash: delta.passwordHash } : {}),
  }
}

/// Outcome of an unprotect attempt. 'unsupported' = the file used a modern
/// algorithmName/hashValue pair this app cannot verify.
export function unprotectPasswordStatus(
  file: Pick<WorkbookSheetProtection, 'hasPassword' | 'passwordHash'>,
  password: string,
): 'ok' | 'wrong' | 'unsupported' {
  if (!file.hasPassword) return 'ok'
  if (file.passwordHash === undefined) return 'unsupported'
  return legacyPasswordMatches(password, file.passwordHash) ? 'ok' : 'wrong'
}

/// True when the file's column-default xf for `screenColumn` says unlocked.
/// Screen → file coordinates go through the journaled structural stream, the
/// same mapping viewport reads use.
function columnDefaultUnlocked(
  state: LazyWorkbookState,
  sheetId: string,
  screenColumn: number,
): boolean {
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (sheet === undefined) return false
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const fileColumn = screenToFile(ops, 'column', screenColumn)
  if (fileColumn === null) return false
  for (const span of sheet.columnWidths) {
    if (fileColumn < span.startColumn || fileColumn > span.endColumn) continue
    if (span.styleIndex === undefined) return false
    return state.file.styles[span.styleIndex]?.locked === false
  }
  return false
}

/// Excel's locked resolution for one cell (see the module doc). `worksheet`
/// is optional so pure callers (tests) can check journal/column rules only.
export function cellIsUnlocked(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
  worksheet?: ProtectionWorksheet,
): boolean {
  const patched = state.editJournal.cells.get(sheetId)?.get(`${row}:${column}`)
    ?.style?.protectionLocked
  if (patched === false) return true
  if (patched === true) return false
  // Raw first: the composed read can lose the flag (see ProtectionWorksheet).
  const custom =
    worksheet?.getCellRaw?.(row, column)?.custom ?? worksheet?.getCell?.(row, column)?.custom
  if (custom?.[UNLOCKED_CELL_KEY] === true) return true
  return columnDefaultUnlocked(state, sheetId, column)
}

/// A set-range-values patch matrix walk result: the first locked cell the
/// patch touches (null when every touched cell is unlocked) plus whether the
/// patch is formatting-only (Excel gates styling on the formatCells
/// attribute rather than the edit prohibition).
export interface PatchInspection {
  readonly locked: { row: number; column: number } | null
  readonly allStyleOnly: boolean
}

/// Walks a set-range-values patch (IObjectMatrix, row → column → ICellData,
/// under either the command's `value` or the mutation's `cellValue` key) and
/// inspects the cells it touches. `formatCellsAllowed` short-circuits the
/// locked scan for formatting-only patches, matching Excel: with formatting
/// allowed, styling a locked cell is legal and the scan only looks at value
/// or formula patches.
export function inspectProtectionPatch(
  state: LazyWorkbookState,
  sheetId: string,
  patchMatrix: unknown,
  formatCellsAllowed: boolean,
  worksheet?: ProtectionWorksheet,
): PatchInspection {
  if (patchMatrix === null || patchMatrix === undefined || typeof patchMatrix !== 'object') {
    return { locked: null, allStyleOnly: false }
  }
  let allStyleOnly = true
  for (const rowKey of Object.keys(patchMatrix as Record<string, unknown>)) {
    const row = Number(rowKey)
    if (!Number.isInteger(row)) return { locked: null, allStyleOnly: false }
    const rowValue = (patchMatrix as Record<string, unknown>)[rowKey]
    if (rowValue === null || rowValue === undefined || typeof rowValue !== 'object') continue
    for (const columnKey of Object.keys(rowValue as Record<string, unknown>)) {
      const column = Number(columnKey)
      if (!Number.isInteger(column)) return { locked: null, allStyleOnly: false }
      const patch = (rowValue as Record<string, unknown>)[columnKey] as
        Record<string, unknown> | null | undefined
      if (patch === null || patch === undefined) continue
      const keys = Object.keys(patch)
      const styleOnly = keys.length > 0 && keys.every((key) => key === 's')
      if (!styleOnly) allStyleOnly = false
      if (styleOnly && formatCellsAllowed) continue
      if (!cellIsUnlocked(state, sheetId, row, column, worksheet)) {
        return { locked: { row, column }, allStyleOnly }
      }
    }
  }
  return { locked: null, allStyleOnly }
}

/// The gate's action families — the protection attributes they consult.
export type ProtectedAction = keyof SheetProtectionAttributes

/// Raw OOXML polarity: true (or absent — the prevented-by-default set) =
/// the action is refused while the sheet is protected. (The two select*
/// attributes default to allowed.)
export function actionPrevented(
  protection: EffectiveSheetProtection,
  action: ProtectedAction,
): boolean {
  const value = protection[action]
  return value !== undefined
    ? value
    : action !== 'selectLockedCells' && action !== 'selectUnlockedCells'
}

/// Minimal read surface the gate needs from a worksheet — the core
/// Worksheet satisfies it, and tests can pass fakes. The raw read is the
/// authoritative one (BUG-1715): composed getCell results go through the
/// CELL_CONTENT interceptor chain, whose rich-text/render handlers rebuild
/// the cell object and can drop the custom bag — in the live app that hid
/// the install flag and refused unlocked cells. The raw cell matrix keeps it.
export interface ProtectionWorksheet {
  getCell?(
    row: number,
    column: number,
  ): { custom?: Record<string, unknown> | null } | null | undefined
  getCellRaw?(
    row: number,
    column: number,
  ): { custom?: Record<string, unknown> | null } | null | undefined
}

/// Bounds of an IRange-shaped command param, when fully present.
export interface ParamBounds {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export function paramRangeBounds(range: unknown): ParamBounds | null {
  if (range === null || range === undefined || typeof range !== 'object') return null
  const candidate = range as Partial<ParamBounds>
  if (
    candidate.startRow === undefined ||
    candidate.endRow === undefined ||
    candidate.startColumn === undefined ||
    candidate.endColumn === undefined
  ) {
    return null
  }
  return {
    startRow: candidate.startRow,
    endRow: candidate.endRow,
    startColumn: candidate.startColumn,
    endColumn: candidate.endColumn,
  }
}

/// Row-major locked scan over a range with a hard cap: a false "all
/// unlocked" on a pathological range only means Excel-parity degrades to a
/// permitted edit, never to data damage (the journal/save stay consistent).
export function rangeHasLockedCell(
  state: LazyWorkbookState,
  sheetId: string,
  bounds: ParamBounds,
  worksheet?: ProtectionWorksheet,
  cap = 4_096,
): boolean {
  let visited = 0
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      if (visited >= cap) return false
      visited += 1
      if (!cellIsUnlocked(state, sheetId, row, column, worksheet)) return true
    }
  }
  return false
}

/// i18n keys the edit gate can raise for a refused command.
export type ProtectionRefusalKey =
  /// The Excel-style "cell you are trying to change is protected" message.
  | 'appSheetCellProtected'
  /// A modeled action (insert/delete rows/columns, sort, autofilter,
  /// row/column formatting) the protection attributes refuse.
  | 'appSheetActionProtected'

const ROW_FORMAT_COMMAND_PATTERN =
  /^sheet\.command\.(set-row-height|delta-row-height|set-row-data|set-rows-hidden|set-specific-rows-visible|set-selected-rows-visible|set-rows-auto-height)$/
const COLUMN_FORMAT_COMMAND_PATTERN =
  /^sheet\.command\.(delta-column-width|set-col-data|set-col-hidden|set-col-visible-on-cols|set-selected-cols-visible)$/

function structuralFamily(eventId: string): ProtectedAction | null {
  if (/^sheet\.command\.(insert-row|insert-multi-rows)/.test(eventId)) return 'insertRows'
  if (/^sheet\.command\.(remove-row)/.test(eventId)) return 'deleteRows'
  if (/^sheet\.command\.(insert-col|insert-multi-cols)/.test(eventId)) return 'insertColumns'
  if (/^sheet\.command\.(remove-col)/.test(eventId)) return 'deleteColumns'
  return null
}

/// The gate's brain: given the effective protection of the target sheet and
/// the pending command, returns the i18n key of the refusal message, or null
/// when the command may proceed. Pure apart from the locked-cell reads.
export function protectionRefusal(
  state: LazyWorkbookState,
  sheetId: string,
  eventId: string,
  params: unknown,
  worksheet: ProtectionWorksheet | undefined,
  protection: EffectiveSheetProtection,
): ProtectionRefusalKey | null {
  // Value / style / paste / clear / autofill writes all land as
  // set-range-values (command or direct mutation).
  if (eventId === SET_RANGE_VALUES_COMMAND || eventId === SET_RANGE_VALUES_MUTATION) {
    const patch =
      (params as { value?: unknown; cellValue?: unknown } | undefined)?.value ??
      (params as { cellValue?: unknown } | undefined)?.cellValue
    const inspection = inspectProtectionPatch(
      state,
      sheetId,
      patch,
      !actionPrevented(protection, 'formatCells'),
      worksheet,
    )
    return inspection.locked === null ? null : 'appSheetCellProtected'
  }
  if (STRUCTURAL_EDIT_COMMAND_PATTERN.test(eventId)) {
    const family = structuralFamily(eventId)
    if (family !== null) {
      return actionPrevented(protection, family) ? 'appSheetActionProtected' : null
    }
    // Merges rewrite the covered cells; Excel refuses them when locked cells
    // are involved.
    const bounds = paramRangeBounds((params as { range?: unknown } | undefined)?.range)
    if (bounds && rangeHasLockedCell(state, sheetId, bounds, worksheet)) {
      return 'appSheetCellProtected'
    }
    return null
  }
  if (SORT_COMMAND_PATTERN.test(eventId) && actionPrevented(protection, 'sort')) {
    return 'appSheetActionProtected'
  }
  if (
    (FILTER_COMMAND_PATTERN.test(eventId) || eventId === OPEN_FILTER_PANEL_OPERATION) &&
    actionPrevented(protection, 'autoFilter')
  ) {
    return 'appSheetActionProtected'
  }
  if (
    eventId === MOVE_RANGE_COMMAND ||
    eventId === MOVE_ROWS_COMMAND ||
    eventId === MOVE_COLS_COMMAND
  ) {
    const rangeParams = params as
      { fromRange?: unknown; toRange?: unknown; range?: unknown } | undefined
    for (const candidate of [rangeParams?.fromRange, rangeParams?.toRange, rangeParams?.range]) {
      const bounds = paramRangeBounds(candidate)
      if (bounds && rangeHasLockedCell(state, sheetId, bounds, worksheet)) {
        return 'appSheetCellProtected'
      }
    }
    return null
  }
  if (ROW_FORMAT_COMMAND_PATTERN.test(eventId) && actionPrevented(protection, 'formatRows')) {
    return 'appSheetActionProtected'
  }
  if (COLUMN_FORMAT_COMMAND_PATTERN.test(eventId) && actionPrevented(protection, 'formatColumns')) {
    return 'appSheetActionProtected'
  }
  return null
}

/**
 * On-demand sheet grid growth (BUG-1615).
 *
 * A new blank workbook used to be pinned to Univer's 1000×26 default grid, so
 * typing or jumping below row 1000 / right of column Z failed with a silent
 * "Range is out of bounds" — while files with millions of rows opened fine.
 * Full Excel-sized grids (1,048,576 × 16,384) stay O(1) to create (the cell
 * matrix is sparse), but the app has dense O(rows × columns) consumers keyed
 * on the grid size — Univer's SetStyleCommand iterates every cell of a styled
 * range (a full-column bold on a 1M-row grid measured 2.5s and a ~1GB heap
 * spike), and the delete→#REF! cross-sheet scan reads whole-sheet
 * getFormulas(). Growing every new book to the Excel maximum would turn
 * ordinary full-column actions into multi-second freezes, so instead the grid
 * grows on demand — Google-Sheets style — when navigation needs room beyond
 * it: the Name Box / Go To jump, and arrow keys (plain or Ctrl) that hit the
 * grid's edge.
 *
 * Growth rides journalSuppression: it is model plumbing, not a user edit, so
 * it must not push an undo entry (Ctrl+Z after "grow, then type" reverts just
 * the typing; the grid stays grown) and must not journal rows — the save's
 * sheet <dimension> is derived from written cells and must stay free of empty
 * tails.
 */
import { Direction, ICommandService } from '@univerjs/core'

import { parseRange } from '../domain/cell-address'
import { journalSuppression, type UniverRuntime, type UniverWorksheet } from './univer-state'

/// Excel's hard sheet limits: rows 1..1048576, columns A..XFD.
export const SHEET_MAX_ROWS = 1_048_576
export const SHEET_MAX_COLUMNS = 16_384

/// Growth steps sized so navigation lands in roomy empty space without ever
/// pre-allocating the Excel maximum (dense range operations iterate the whole
/// grid — see the module comment).
const ROW_GROWTH_STEP = 5_000
const COLUMN_GROWTH_STEP = 64

const MOVE_SELECTION_COMMAND = 'sheet.command.move-selection'

/// Sizes a single axis: a minimum at or below the current count (0 = don't
/// care) never grows; otherwise the count rises past the minimum by a growth
/// step, capped at the Excel limit.
function grownCount(current: number, minNeeded: number, step: number, cap: number): number {
  if (minNeeded <= current) return current
  return Math.min(Math.max(minNeeded, current + step), cap)
}

/**
 * Grows the sheet's grid so it holds at least `minRows` × `minColumns` rows
 * and columns (exclusive 0-based counts: the row with index `minRows - 1`
 * exists after the call; 0 = don't care). Appends empty rows/columns at the
 * end through Univer's insert mutations — no existing cell moves, and
 * nothing is materialized beyond the empty axis records. Returns the
 * (possibly unchanged) grid counts.
 */
export function ensureSheetGrid(
  runtime: UniverRuntime,
  workbook: GrowthWorkbook,
  worksheet: UniverWorksheet,
  minRows: number,
  minColumns: number,
): { rows: number; columns: number } {
  const beforeRows = worksheet.getMaxRows()
  const beforeColumns = worksheet.getMaxColumns()
  const rows = grownCount(beforeRows, minRows, ROW_GROWTH_STEP, SHEET_MAX_ROWS)
  const columns = grownCount(beforeColumns, minColumns, COLUMN_GROWTH_STEP, SHEET_MAX_COLUMNS)
  if (rows === beforeRows && columns === beforeColumns) return { rows, columns }
  const commandService = runtime.univer.__getInjector().get(ICommandService)
  const params = { unitId: workbook.getId(), subUnitId: worksheet.getSheetId() }
  // journalSuppression drops the mutations' undo entries (the prototype-patch
  // undo filter must be installed — App does at startup) and keeps the edit
  // journal from replaying growth as structural ops at save time.
  const restoreSuppression = journalSuppression.active
  journalSuppression.active = true
  try {
    if (columns !== beforeColumns) {
      commandService.syncExecuteCommand('sheet.mutation.insert-col', {
        ...params,
        range: { startRow: 0, endRow: 0, startColumn: beforeColumns, endColumn: columns - 1 },
      })
    }
    if (rows !== beforeRows) {
      commandService.syncExecuteCommand('sheet.mutation.insert-row', {
        ...params,
        range: { startRow: beforeRows, endRow: rows - 1, startColumn: 0, endColumn: 0 },
      })
    }
  } finally {
    journalSuppression.active = restoreSuppression
  }
  return { rows: worksheet.getMaxRows(), columns: worksheet.getMaxColumns() }
}

/// Minimal structural view of the facade workbook the growth helpers need
/// (FWorkbook satisfies it; the shape keeps the module unit-testable).
export interface GrowthWorkbook {
  getId(): string
  getActiveSheet(): UniverWorksheet | null
  getSheetByName(name: string): UniverWorksheet | null
  getActiveRange(): {
    getRow(): number
    getColumn(): number
    getHeight(): number
    getWidth(): number
  } | null
}

/// Strips an optional `Sheet1!` / `'My Sheet'!` prefix, resolving the sheet
/// against the workbook; null when the prefix names no sheet of this
/// workbook (the caller keeps its existing unknown-sheet error).
function sheetForRef(
  workbook: GrowthWorkbook,
  ref: string,
): { worksheet: UniverWorksheet; body: string } | null {
  const bang = ref.lastIndexOf('!')
  if (bang === -1) {
    const worksheet = workbook.getActiveSheet()
    return worksheet ? { worksheet, body: ref } : null
  }
  const rawName = ref.slice(0, bang)
  const name = rawName.replace(/^'/, '').replace(/'$/, '')
  const worksheet = workbook.getSheetByName(name)
  return worksheet ? { worksheet, body: ref.slice(bang + 1) } : null
}

/**
 * Grows the referenced sheet's grid to fit a Name Box / Go To target — an A1
 * address, range, whole-column span ("A:C"), or whole-row span ("3:8"),
 * optionally sheet-qualified. Out-of-range endpoints beyond the Excel limits
 * were already rejected upstream (goto.ts bounds checks) and simply clamp.
 * Returns null when the ref names an unknown sheet (the jump then fails as
 * before); otherwise the target worksheet.
 */
export function growSheetToRef(runtime: UniverRuntime, ref: string): UniverWorksheet | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return null
  const resolved = sheetForRef(workbook, ref)
  if (!resolved) return null
  const { worksheet, body } = resolved
  let minRows = 0
  let minColumns = 0
  for (const part of body.replace(/\$/g, '').split(':')) {
    if (/^\d+$/.test(part)) {
      // whole-row span ("1500:1500"): the row number is 1-based
      minRows = Math.max(minRows, Number(part))
      continue
    }
    if (/^[A-Za-z]{1,3}$/.test(part)) {
      // whole-column span ("AB:AB"): the column label is 1-based
      minColumns = Math.max(minColumns, parseRange(`${part}1`).startColumn + 1)
      continue
    }
    const cell = parseRange(part)
    minRows = Math.max(minRows, cell.endRow + 1)
    minColumns = Math.max(minColumns, cell.endColumn + 1)
  }
  ensureSheetGrid(runtime, workbook, worksheet, minRows, minColumns)
  return worksheet
}

/**
 * Arrow-key edge growth, one command dispatch: a plain or Ctrl arrow that
 * would move the selection past the grid's last row/column grows the grid by
 * a step first, so keyboard navigation beyond a new book's 1000×26 works
 * like on an imported million-row file. UP/LEFT never grow; mid-sheet moves
 * never grow (ensureSheetGrid only reacts when the minimum exceeds the
 * current size). Split from the listener so tests can drive it without
 * Univer's sheets-ui command registry.
 */
export function growOnArrowAtEdge(
  runtime: UniverRuntime,
  commandId: string,
  direction: Direction,
): void {
  if (commandId !== MOVE_SELECTION_COMMAND) return
  if (direction !== Direction.DOWN && direction !== Direction.RIGHT) return
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return
  // Grow around the active selection; getActiveRange can fail while no
  // selection exists and is then skipped.
  const range = (() => {
    try {
      return workbook.getActiveRange()
    } catch {
      return null
    }
  })()
  if (!range) return
  ensureSheetGrid(
    runtime,
    workbook,
    worksheet,
    direction === Direction.DOWN ? range.getRow() + range.getHeight() + 1 : 0,
    direction === Direction.RIGHT ? range.getColumn() + range.getWidth() + 1 : 0,
  )
}

/**
 * Installs arrow-key edge growth for the session (before-command gate).
 */
export function installEdgeNavigationGrowth(runtime: UniverRuntime): { dispose(): void } {
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
    const params = event.params as { direction?: Direction } | undefined
    growOnArrowAtEdge(runtime, event.id, params?.direction ?? Direction.UP)
  })
}

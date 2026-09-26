/**
 * Table Design support (PAR-202): locates the table under the active cell
 * (a session-added table or one already stored in the file), and applies
 * the dialog's changes to the edit journal. File-table edits are refused on
 * sheets with pending row/column shifts — their coordinates are final at
 * record time and the gateway fails closed on any other order.
 */
import { columnLabel, parseRange, type RangeBounds } from '../domain/cell-address'
import { tableNameError } from '../domain/table-refs'
import type { WorkbookFile } from '../shared/desktop-api'
import {
  recordTableEdit,
  removeSlicerAddsForTable,
  removeTableAdd,
  updateTableAdd,
  type EditJournal,
} from './edit-journal'
import { t } from './i18n/locale'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

export type FileTable = WorkbookFile['sheets'][number]['tables'][number]

export interface TableDesignSeed {
  readonly origin: 'session' | 'file'
  readonly sheetId: string
  /// The address edits use: the journal name for session tables, the file's
  /// displayName token for file tables.
  readonly tableName: string
  readonly name: string
  readonly area: RangeBounds
  readonly columnNames: readonly string[]
  readonly style: string | undefined
  readonly bandedRows: boolean
  /// Sidecar-resolved row-stripe color; Convert to Range bakes it into the
  /// body cells because the style's banding dies with the table part.
  readonly stripeFill: string | undefined
  readonly tableId: string | undefined
}

export type TableDesignTarget =
  | { readonly kind: 'none'; readonly message: string }
  | { readonly kind: 'table'; readonly seed: TableDesignSeed }

function hasPendingShifts(journal: EditJournal, sheetId: string): boolean {
  return (journal.structuralOps.get(sheetId)?.length ?? 0) > 0
}

/// Finds the table covering (row, column): session-added tables first, then
/// file tables adjusted by any pending file-table edit.
export function findTableDesignTarget(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): TableDesignTarget {
  const journal = state.editJournal
  const contains = (area: RangeBounds): boolean =>
    row >= area.startRow &&
    row <= area.endRow &&
    column >= area.startColumn &&
    column <= area.endColumn

  const session = journal.tableAdds.find(
    (table) => table.sheetId === sheetId && contains(table.area),
  )
  if (session) {
    return {
      kind: 'table',
      seed: {
        origin: 'session',
        sheetId,
        tableName: session.name,
        name: session.name,
        area: session.area,
        columnNames: session.columnNames,
        style: session.style,
        bandedRows: session.bandedRows,
        // Session banding already lives in the cells (applyTableBanding at
        // load), so nothing needs baking when the table goes away.
        stripeFill: undefined,
        tableId: session.tableId,
      },
    }
  }

  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) return { kind: 'none', message: t('dlgTableNotInTable') }
  const table = sheet.tables.find((candidate) => contains(candidate.range))
  if (!table) return { kind: 'none', message: t('dlgTableNotInTable') }
  const entry = journal.tableEdits.find(
    (candidate) =>
      candidate.sheetId === sheetId &&
      candidate.tableName.toLowerCase() === (table.name ?? '').toLowerCase(),
  )
  if (entry?.convertToRange === true) {
    return { kind: 'none', message: t('dlgTableNotInTable') }
  }
  const area = entry?.resize?.area ?? table.range
  return {
    kind: 'table',
    seed: {
      origin: 'file',
      sheetId,
      tableName: table.name ?? '',
      name: entry?.rename ?? table.name ?? '',
      area,
      columnNames: table.columns ?? [],
      style: entry?.style?.style ?? table.styleName,
      bandedRows: entry?.style?.bandedRows ?? table.showRowStripes,
      stripeFill: table.stripeFill,
      tableId: undefined,
    },
  }
}

/// True when the name would collide with another table or a defined name.
function nameTaken(
  state: LazyWorkbookState,
  name: string,
  keepOrigin: {
    readonly origin: 'session' | 'file'
    readonly tableName: string
  },
): boolean {
  const needle = name.toLowerCase()
  for (const table of state.editJournal.tableAdds) {
    if (
      keepOrigin.origin === 'session' &&
      table.name.toLowerCase() === keepOrigin.tableName.toLowerCase()
    ) {
      continue
    }
    if (table.name.toLowerCase() === needle) return true
  }
  const file = state.file.sheets.flatMap((sheet) => sheet.tables)
  for (const table of file) {
    if (
      keepOrigin.origin === 'file' &&
      (table.name ?? '').toLowerCase() === keepOrigin.tableName.toLowerCase()
    ) {
      continue
    }
    if ((table.name ?? '').toLowerCase() === needle) return true
  }
  for (const name2 of state.file.definedNames) {
    if (name2.name.toLowerCase() === needle) return true
  }
  return false
}

/// Appends Excel-style unique names for grown columns.
function growColumnNames(existing: readonly string[], width: number): string[] {
  const names = [...existing]
  if (names.length === 0) return names
  const used = new Set(names.map((name) => name.trim().toLowerCase()))
  let nextId = names.length + 1
  while (names.length < width) {
    let candidate = `Column${nextId}`
    for (let suffix = 2; used.has(candidate.toLowerCase()); suffix += 1) {
      candidate = `Column${nextId}_${suffix}`
    }
    used.add(candidate.toLowerCase())
    names.push(candidate)
    nextId += 1
  }
  return names
}

export interface TableDesignChange {
  readonly seed: TableDesignSeed
  readonly name: string
  readonly area: RangeBounds
  readonly style: string | undefined
  readonly bandedRows: boolean
  readonly convertToRange: boolean
}

/// Applies the dialog's change; returns a localized error message or null.
export function applyTableDesignChange(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
  change: TableDesignChange,
): string | null {
  const { seed } = change
  const journal = state.editJournal
  if (change.convertToRange) {
    // Excel removes a table's slicers when the table converts to a range —
    // the slicer cache has no source without the table part (whose autoFilter
    // and criteria die with it). Drop the journal entries so the save does
    // not try to bind slicers to a part that no longer exists; the App
    // removes the matching panels the same way.
    removeSlicerAddsForTable(journal, seed.sheetId, seed.tableName)
    if (seed.origin === 'session') {
      const removed = removeTableAdd(journal, seed.sheetId, seed.tableName)
      if (!removed) return t('appCommandFailed')
      const sheet = runtime?.univerAPI.getActiveWorkbook()?.getSheetBySheetId(seed.sheetId)
      if (seed.tableId !== undefined && sheet) {
        fireAndForget(sheet.removeTable(seed.tableId))
      }
      return null
    }
    if (hasPendingShifts(journal, seed.sheetId)) return t('dlgTablePendingShifts')
    recordTableEdit(journal, {
      sheetId: seed.sheetId,
      tableName: seed.tableName,
      convertToRange: true,
      // The dialog's final banding state decides whether the stripes are
      // worth baking; the color is the sidecar-resolved style stripe.
      ...(change.bandedRows && seed.stripeFill ? { stripeFill: seed.stripeFill } : {}),
    })
    return null
  }

  const nameError = tableNameError(change.name)
  if (nameError !== null) return t('dlgTableBadName')
  if (
    change.name.toLowerCase() !== seed.tableName.toLowerCase() &&
    nameTaken(state, change.name, { origin: seed.origin, tableName: seed.tableName })
  ) {
    return t('dlgTableNameTaken')
  }
  if (
    change.area.endRow <= change.area.startRow ||
    change.area.endColumn < change.area.startColumn
  ) {
    return t('dlgTableBadRange')
  }

  if (seed.origin === 'session') {
    const areaChanged =
      change.area.startRow !== seed.area.startRow ||
      change.area.startColumn !== seed.area.startColumn ||
      change.area.endRow !== seed.area.endRow ||
      change.area.endColumn !== seed.area.endColumn
    if (areaChanged) {
      if (
        change.area.startRow !== seed.area.startRow ||
        change.area.startColumn !== seed.area.startColumn
      ) {
        return t('dlgTableMoveHeader')
      }
    }
    const columnNames =
      change.area.endColumn - change.area.startColumn + 1 === seed.columnNames.length
        ? seed.columnNames
        : growColumnNames(seed.columnNames, change.area.endColumn - change.area.startColumn + 1)
    const updated = updateTableAdd(journal, seed.sheetId, seed.tableName, {
      ...(change.name !== seed.tableName ? { name: change.name } : {}),
      ...(areaChanged ? { area: change.area } : {}),
      ...(areaChanged ? { columnNames } : {}),
      ...(change.style !== seed.style ? { style: change.style } : {}),
      ...(change.bandedRows !== seed.bandedRows ? { bandedRows: change.bandedRows } : {}),
    })
    if (!updated) return t('appCommandFailed')
    if (areaChanged && seed.tableId !== undefined) {
      resyncSessionTable(runtime, seed, change)
    }
    return null
  }

  if (hasPendingShifts(journal, seed.sheetId)) return t('dlgTablePendingShifts')
  if (
    change.area.startRow !== seed.area.startRow ||
    change.area.startColumn !== seed.area.startColumn
  ) {
    return t('dlgTableMoveHeader')
  }
  const currentName = seed.name
  recordTableEdit(journal, {
    sheetId: seed.sheetId,
    tableName: seed.tableName,
    ...(change.name !== currentName ? { rename: change.name } : {}),
    ...(change.area.startRow !== seed.area.startRow ||
    change.area.startColumn !== seed.area.startColumn ||
    change.area.endRow !== seed.area.endRow ||
    change.area.endColumn !== seed.area.endColumn
      ? { resize: { area: change.area } }
      : {}),
    ...(change.style !== seed.style || change.bandedRows !== seed.bandedRows
      ? {
          style: {
            ...(change.style !== seed.style ? { style: change.style } : {}),
            ...(change.bandedRows !== seed.bandedRows ? { bandedRows: change.bandedRows } : {}),
          },
        }
      : {}),
  })
  return null
}

/// Re-registers a resized session table in Univer (banding/filter overlay).
/// Cosmetic and best-effort: the journal entry is what the save writes.
function resyncSessionTable(
  runtime: UniverRuntime | null,
  seed: TableDesignSeed,
  change: TableDesignChange,
): void {
  if (!runtime || seed.tableId === undefined) return
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(seed.sheetId)
  if (!worksheet) return
  fireAndForget(worksheet.removeTable(seed.tableId), () => {
    fireAndForget(
      worksheet.addTable(
        change.name,
        {
          startRow: change.area.startRow,
          startColumn: change.area.startColumn,
          endRow: change.area.endRow,
          endColumn: change.area.endColumn,
        },
        seed.tableId!,
      ),
    )
  })
}

/// The Univer table facade mixes sync and async command overloads; swallow
/// either shape's failures — these calls are cosmetic only.
function fireAndForget(result: unknown, after?: () => void): void {
  try {
    Promise.resolve(result)
      .then(() => after?.())
      .catch(() => undefined)
  } catch {
    // Sync rejection: nothing to resync.
  }
}

/// A1 label of an area for the dialog's range field.
export function areaToRangeText(area: RangeBounds): string {
  return `${columnLabel(area.startColumn)}${area.startRow + 1}:${columnLabel(area.endColumn)}${
    area.endRow + 1
  }`
}

/// Parses the dialog's range field; null when malformed.
export function parseRangeText(text: string): RangeBounds | null {
  try {
    return parseRange(text.trim())
  } catch {
    return null
  }
}

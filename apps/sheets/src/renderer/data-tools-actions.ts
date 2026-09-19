/**
 * Cell navigation (Name Box / Go To / formula / symbol) and data tools
 * (advanced filter, subtotal, consolidate, outline, format-as-table).
 * Extracted from App.tsx; the App component passes a DataToolsContext built
 * fresh per call so refs and state never go stale.
 */
import { ILayoutService } from '@univerjs/preset-sheets-core'

import { legacyCharsetForLang } from '@airy-office/file-parse/text'

import { columnLabel } from '../domain/cell-address'
import type { AdvancedFilterColumn, AdvancedFilterCriteria } from './AdvancedFilterDialog'
import {
  buildLabelMatrix,
  buildPositionMatrix,
  parseConsolidateReference,
  targetOverlapsSource,
  type ConsolidateArea,
  type ConsolidateConfig,
  type OutputCell,
} from './consolidate'
import { isSheetRemoved, journalSize, recordPageSetup, recordStructuralOp } from './edit-journal'
import { resolveGoToRef, type GoToNameEntry } from './goto'
import { getLang, t } from './i18n/locale'
import { appendSymbol } from './SymbolDialog'
import {
  coerceFieldValue,
  DATE_COLUMN_TYPES,
  parseDestinationCell,
  splitDelimited,
  splitFixedWidth,
  activeDelimiterChars,
  type TextToColumnsConfig,
} from './text-to-columns'
import {
  a1RangeRef,
  a1RowRangeRef,
  advancedFilterColumnOptions,
  applyFilterCriteria,
  columnLetter,
  loadVisibleRange,
  sheetOutline,
  univerDefinedNames,
  revealCellBelowFreeze,
} from './univer-sync'
import { pushVisualUndo } from './univer-sync'
import { outlineUndoGate } from './univer-state'
import { MAX_OUTLINE_LEVEL, placementForAxis, type OutlinePlacement } from './outline'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import { applyAiTableAdd } from './workbook-ops'

/** The App refs/state the data-tool actions need; built fresh per call. */
export interface DataToolsContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  setAdvancedFilterColumns: (columns: readonly AdvancedFilterColumn[] | null) => void
  /// Outline edits notify the gutter so it re-renders its +/- buttons.
  onOutlineChanged?: (() => void) | undefined
}

/// Matches the paste ceiling in spirit: one setValues command, journaled and
/// undoable; larger files should open as their own workbook instead.
const CSV_IMPORT_MAX_CELLS = 50_000

/// PERF-902: the CSV pipeline drags jszip in for its xlsx-conversion output
/// (~144 kB, the only eager jszip importer). It is needed solely on an
/// explicit Data → From Text/CSV action, so it loads on demand instead of
/// riding the eager vendor-misc chunk.
type CsvImport = typeof import('../gateway/csv-import')

/// Data → From Text/CSV: reads a delimited file (encoding and delimiter
/// sniffed by the shared CSV pipeline) into the active sheet at the selection.
export function handleImportCsv(ctx: DataToolsContext): void {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = '.csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain'
  input.onchange = () => {
    const file = input.files?.[0]
    if (!file) return
    if (file.size > 32 * 1024 * 1024) {
      ctx.setMessage(t('appCsvTooLarge'))
      return
    }
    void importCsvFile(ctx, file)
  }
  input.click()
}

/// BUG-1201: the lazy csv-import chunk and the legacy-charset decode run
/// before importCsvText (whose try/catch only covers setValues), so after
/// PERF-902 a broken chunk or a decoding failure died as an unhandled
/// rejection while the user's picked file silently did nothing. One boundary
/// here reports it through the message bar instead — the same string the
/// setValues path fails with.
export async function importCsvFile(ctx: DataToolsContext, file: File): Promise<void> {
  try {
    const buffer = new Uint8Array(await file.arrayBuffer())
    const csv = await import('../gateway/csv-import')
    importCsvText(ctx, csv.decodeCsvBuffer(buffer, legacyCharsetForLang(getLang())), csv)
  } catch {
    // Chunk-load and decode internals are never user-actionable detail.
    ctx.setMessage(t('appCsvImportFailed'))
  }
}

function importCsvText(ctx: DataToolsContext, text: string, csv: CsvImport): void {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const rows = csv.parseCsv(text)
  const columns = rows.reduce((width, row) => Math.max(width, row.length), 0)
  if (rows.length === 0 || columns === 0) {
    ctx.setMessage(t('appCsvEmpty'))
    return
  }
  if (rows.length * columns > CSV_IMPORT_MAX_CELLS) {
    ctx.setMessage(t('appCsvTooLarge'))
    return
  }
  const values = rows.map((row) =>
    Array.from({ length: columns }, (_, index) => {
      const cell = row[index] ?? ''
      return csv.isNumericCell(cell) ? { v: Number(cell) } : { v: cell }
    }),
  )
  const row = range.getRow()
  const column = range.getColumn()
  try {
    worksheet.getRange(row, column, rows.length, columns).setValues(values)
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appCsvImportFailed'))
    return
  }
  ctx.setMessage(
    t('appCsvImported', {
      rows: rows.length,
      columns,
      cell: `${columnLetter(column)}${row + 1}`,
    }),
  )
}

export function activeCellLabel(ctx: DataToolsContext): string {
  const range = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range) return t('appActiveCellFallback')
  return `${columnLetter(range.getColumn())}${range.getRow() + 1}`
}

/// Name Box / Go To jump: resolves the typed reference (an A1 address or a
/// defined name) and selects it, switching sheets when the target lives on
/// another one. Returns null on success or a user-facing error message.
/// Univer's parser yields NaN rows instead of throwing on garbage, so
/// validity is decided by resolveGoToRef — the try/catch covers what
/// getRange itself rejects: unknown sheet names and out-of-bounds ranges.
export function goToReference(ctx: DataToolsContext, ref: string): string | null {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return t('appGoToNotReady')
  const trimmed = ref.trim()
  if (trimmed === '') return t('appGoToEmpty')
  const resolved = resolveGoToRef(trimmed, listDefinedNames(ctx))
  if (resolved === null) {
    return t('appGoToUnresolved', { ref: trimmed })
  }
  try {
    // A jump must not leave an editor open on the previous cell — later
    // keystrokes would land there. Commit it before moving.
    const editing = workbook as unknown as {
      isCellEditing?(): boolean
      endEditingAsync?(save: boolean): Promise<boolean>
    }
    if (editing.isCellEditing?.()) void editing.endEditingAsync?.(true)
    const range = worksheet.getRange(resolved)
    const target = workbook.getSheetBySheetId(range.getSheetId()) ?? worksheet
    workbook.setActiveRange(range)
    void revealCellBelowFreeze(target, range.getRow(), range.getColumn())
    // Hand keyboard focus back to the grid (Univer's hidden editor host, the
    // same handoff its own name box does); typing right after a jump then
    // lands in the target cell instead of being dropped on <body>.
    ctx.univerRef.current?.univer.__getInjector().get(ILayoutService).focus()
    // The single Scroll event a jump produces computes its viewport from a
    // stale getVisibleRange; load around the target explicitly.
    if (ctx.univerRef.current) {
      void loadVisibleRange(
        ctx.univerRef.current,
        ctx.lazyWorkbookRef as { current: LazyWorkbookState | null },
        target,
        ctx.setMessage,
        { row: range.getRow(), column: range.getColumn() },
      )
    }
  } catch (error: unknown) {
    const detail = error instanceof Error ? error.message : ''
    if (detail.includes('out of bounds')) return t('appGoToOutOfBounds', { ref: trimmed })
    if (detail.includes('Range not found')) return t('appGoToSheetNotFound', { ref: trimmed })
    return detail === ''
      ? t('appGoToInvalidAddress', { ref: trimmed })
      : t('appGoToFailed', { ref: trimmed, detail })
  }
  return null
}

/// All defined names of the active workbook, feeding both the Go To
/// dialog's list and Name Box resolution.
export function listDefinedNames(ctx: DataToolsContext): GoToNameEntry[] {
  return univerDefinedNames(ctx.univerRef.current).map((defined) => ({
    name: defined.getName(),
    ref: defined.getFormulaOrRefString(),
  }))
}

export function handleApplyFormula(ctx: DataToolsContext, formula: string): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectCellFirst')
  const trimmed = formula.trim()
  if (!trimmed.startsWith('=')) return t('appFormulaStartsEquals')
  const opens = (trimmed.match(/\(/g) ?? []).length
  const closes = (trimmed.match(/\)/g) ?? []).length
  if (opens !== closes) return t('appUnbalancedParens')
  try {
    worksheet.getRange(range.getRow(), range.getColumn(), 1, 1).setValue({ f: trimmed })
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appSetFormulaFailed')
  }
  ctx.setMessage(t('appFormulaSet', { cell: activeCellLabel(ctx) }))
  return null
}

/// A label turned into a legal defined name, Excel-style: illegal characters
/// become underscores, and anything that could read as a cell reference gets
/// an underscore prefix.
function definedNameFromLabel(label: string): string | null {
  const cleaned = label.trim().replace(/[^\p{L}\p{N}_.]/gu, '_')
  if (!cleaned || /^\.+$/.test(cleaned)) return null
  const named = /^[\p{L}_]/u.test(cleaned) ? cleaned : `_${cleaned}`
  const cellLike = /^[A-Za-z]{1,3}\d+$/.test(named) || /^[Rr]\d*([Cc]\d*)?$/.test(named)
  return cellLike ? `_${named}` : named
}

/// Formulas → Create from Selection: one defined name per column (labels in
/// the selection's top row) or per row (labels in its left column). The
/// defined-name mutation listener journals each insert.
export function handleCreateNamesFromSelection(ctx: DataToolsContext, mode: 'top' | 'left'): void {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const selection = workbook?.getActiveRange()?.getRange()
  if (!workbook || !worksheet || !selection) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const { startRow, endRow, startColumn, endColumn } = selection
  if (mode === 'top' ? endRow <= startRow : endColumn <= startColumn) {
    ctx.setMessage(t('appCreateNamesNeedsData'))
    return
  }
  const sheetName = worksheet.getSheetName()
  // Names span the full selection like Excel (empty data rows included),
  // but labels past the data extent are blank and would be skipped one by
  // one — stop the label walk there so a whole-column selection doesn't
  // read a million empty cells. File cells stream into Univer lazily, so
  // getLastRow/getLastColumn alone can undercount; the file's used range is
  // the floor.
  const fileSheet = ctx.lazyWorkbookRef.current?.file.sheets.find(
    (sheet) => sheet.id === worksheet.getSheetId(),
  )
  const labelEndRow = Math.min(
    endRow,
    Math.max(worksheet.getLastRow(), (fileSheet?.rowCount ?? 0) - 1),
  )
  const labelEndColumn = Math.min(
    endColumn,
    Math.max(worksheet.getLastColumn(), (fileSheet?.columnCount ?? 0) - 1),
  )
  // Labels come from the formatted display text — a date header must name
  // like "Jan_2023", not its raw serial.
  const entries: { label: string; ref: string }[] = []
  if (mode === 'top') {
    for (let column = startColumn; column <= labelEndColumn; column += 1) {
      entries.push({
        label: worksheet.getRange(startRow, column, 1, 1).getDisplayValue() ?? '',
        ref: a1RangeRef(sheetName, column, startRow + 1, endRow),
      })
    }
  } else {
    for (let row = startRow; row <= labelEndRow; row += 1) {
      entries.push({
        label: worksheet.getRange(row, startColumn, 1, 1).getDisplayValue() ?? '',
        ref: a1RowRangeRef(sheetName, row, startColumn + 1, endColumn),
      })
    }
  }
  // insertDefinedName keys entries by internal id and accepts duplicates
  // without error; a second entry with the same name fails the next save.
  // Dedupe here — against existing names and within the batch, both
  // case-insensitive like Excel.
  const taken = new Set(
    univerDefinedNames(ctx.univerRef.current).map((defined) => defined.getName().toLowerCase()),
  )
  let created = 0
  let skipped = 0
  for (const entry of entries) {
    const name = definedNameFromLabel(entry.label)
    if (!name || taken.has(name.toLowerCase())) {
      skipped += 1
      continue
    }
    try {
      workbook.insertDefinedName(name, entry.ref)
      taken.add(name.toLowerCase())
      created += 1
    } catch {
      // Reserved or otherwise rejected name — Excel prompts here; we skip.
      skipped += 1
    }
  }
  ctx.setMessage(
    skipped > 0
      ? t('appNamesCreatedSkipped', { count: created, skipped })
      : t('appNamesCreated', { count: created }),
  )
}

/// The Symbol dialog's click: appends the picked character to the active
/// range's top-left cell as text. setValue lands as a normal
/// set-range-values command, so it journals and undoes like typing.
export function handleInsertSymbol(ctx: DataToolsContext, char: string): void {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSymbolNeedsCell'))
    return
  }
  const cell = worksheet.getRange(range.getRow(), range.getColumn(), 1, 1)
  try {
    cell.setValue(appendSymbol(cell.getValue(), char))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appSymbolInsertFailed'))
    return
  }
  ctx.setMessage(t('appSymbolInserted', { char, cell: activeCellLabel(ctx) }))
}

/// Samples the active filter range's header row and opens the Advanced
/// Filter dialog; without an auto-filter it explains instead of opening.
export function openAdvancedFilterDialog(ctx: DataToolsContext): void {
  const worksheet = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const filter = worksheet?.getFilter()
  if (!worksheet || !filter) {
    ctx.setMessage(t('appAdvFilterNeedsFilter'))
    return
  }
  ctx.setAdvancedFilterColumns(advancedFilterColumnOptions(worksheet, filter))
}

/// The Advanced Filter dialog's OK: lands through the same
/// applyFilterCriteria path as the AI op set_filter_criteria.
export function handleApplyAdvancedFilter(
  ctx: DataToolsContext,
  criteria: AdvancedFilterCriteria,
): string | null {
  const worksheet = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const filter = worksheet?.getFilter()
  if (!worksheet || !filter) return t('appAdvFilterGone')
  const startColumn = filter.getRange().getRange().startColumn
  try {
    applyFilterCriteria(worksheet, columnLabel(startColumn + criteria.colId), {
      customs: { and: criteria.and, filters: criteria.filters },
    })
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appAdvFilterFailed')
  }
  ctx.setMessage(t('appAdvFilterApplied'))
  return null
}

export function handleCreateSubtotal(
  ctx: DataToolsContext,
  config: {
    groupCol: number
    valueCol: number
    agg: 'sum' | 'count' | 'average'
  },
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectSourceRangeFirst')
  const values = range.getValues()
  if (values.length < 2) return t('appSelectRangeWithHeader')
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const groupOffset = config.groupCol - startColumn
  const valueOffset = config.valueCol - startColumn
  if (groupOffset === valueOffset) return t('appSubtotalSameColumns')
  const body = values.slice(1)
  if (body.length === 0) return t('appRangeNoDataRows')
  // SUBTOTAL function numbers: 9 = SUM, 3 = COUNTA, 1 = AVERAGE.
  const fn = config.agg === 'sum' ? 9 : config.agg === 'count' ? 3 : 1
  const valueColumn = columnLetter(config.valueCol)
  const groups: { label: string; startAbs: number; endAbs: number }[] = []
  let currentLabel = String(body[0]?.[groupOffset] ?? '')
  let groupStart = startRow + 1
  for (let index = 1; index <= body.length; index++) {
    const label = index < body.length ? String(body[index]?.[groupOffset] ?? '') : null
    if (label !== currentLabel) {
      groups.push({ label: currentLabel, startAbs: groupStart, endAbs: startRow + index })
      if (label !== null) {
        currentLabel = label
        groupStart = startRow + 1 + index
      }
    }
  }
  if (groups.length > 200) return t('appSubtotalTooManyGroups')
  const writeTotalRow = (row: number, label: string, from: number, to: number): void => {
    const labelCell = worksheet.getRange(row, config.groupCol, 1, 1)
    labelCell.setValue({ v: label })
    labelCell.setFontWeight('bold')
    const valueCell = worksheet.getRange(row, config.valueCol, 1, 1)
    valueCell.setValue({
      f: `=SUBTOTAL(${fn},$${valueColumn}$${from + 1}:$${valueColumn}$${to + 1})`,
    })
    valueCell.setFontWeight('bold')
  }
  try {
    // Bottom-up so earlier group indices stay valid across the inserts.
    for (const group of [...groups].reverse()) {
      worksheet.insertRowsBefore(group.endAbs + 1, 1)
      writeTotalRow(group.endAbs + 1, `${group.label} Total`, group.startAbs, group.endAbs)
    }
    const lastRowAbs = startRow + body.length + groups.length
    worksheet.insertRowsBefore(lastRowAbs + 1, 1)
    // SUBTOTAL skips the nested per-group subtotal rows inside the span.
    writeTotalRow(lastRowAbs + 1, 'Grand Total', startRow + 1, lastRowAbs)
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appSubtotalInsertFailed')
  }
  ctx.setMessage(t('appSubtotalsInserted', { count: groups.length }))
  return null
}

export function consolidateDefaultReference(ctx: DataToolsContext): string {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range || (range.getHeight() === 1 && range.getWidth() === 1)) return ''
  const name = worksheet.getSheetName()
  const prefix = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`
  return `${prefix}!${range.getA1Notation()}`
}

export function handleCreateConsolidate(
  ctx: DataToolsContext,
  config: ConsolidateConfig,
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!workbook || !worksheet || !active) return t('appSelectTargetCellFirst')
  const activeSheetName = worksheet.getSheetName()
  const sheets = workbook.getSheets()
  const areas: ConsolidateArea[] = []
  const sources: UniverWorksheet[] = []
  for (const reference of config.references) {
    const parsed = parseConsolidateReference(reference)
    if (!parsed) return t('appConsolidateBadRef', { ref: reference })
    const name = parsed.sheetName ?? activeSheetName
    const source = sheets.find((sheet) => sheet.getSheetName().toLowerCase() === name.toLowerCase())
    if (!source) return t('appNoSheetNamed', { name })
    sources.push(source)
    areas.push({
      sheetName: source.getSheetName(),
      startRow: parsed.range.startRow,
      startColumn: parsed.range.startColumn,
      rows: parsed.range.endRow - parsed.range.startRow + 1,
      columns: parsed.range.endColumn - parsed.range.startColumn + 1,
    })
  }
  const target = { row: active.getRow(), column: active.getColumn() }
  let matrix: OutputCell[][]
  if (config.leftLabels) {
    const areaLabels = areas.map((area, index) =>
      (
        sources[index]?.getRange(area.startRow, area.startColumn, area.rows, 1).getValues() ?? []
      ).map((row) => String(row?.[0] ?? '')),
    )
    const built = buildLabelMatrix(config.fn, areas, areaLabels, target)
    if (typeof built === 'string') return built
    matrix = built
  } else {
    matrix = buildPositionMatrix(config.fn, areas)
  }
  const columns = matrix[0]?.length ?? 0
  if (matrix.length === 0 || columns === 0) return t('appConsolidateEmptySources')
  if (
    targetOverlapsSource(areas, activeSheetName, {
      ...target,
      rows: matrix.length,
      columns,
    })
  ) {
    return t('appConsolidateOverlap')
  }
  try {
    worksheet
      .getRange(target.row, target.column, matrix.length, columns)
      .setValues(
        matrix.map((line) =>
          line.map((cell) => (cell.f !== undefined ? { f: cell.f } : { v: cell.v ?? null })),
        ),
      )
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appConsolidateWriteFailed')
  }
  ctx.setMessage(
    t('appConsolidateDone', {
      count: areas.length,
      cell: `${columnLetter(target.column)}${target.row + 1}`,
      rows: matrix.length,
      columns,
    }),
  )
  return null
}

/// The sheet's outline summary placement: the session journal over the
/// file's outlinePr (unknown file values fall back to Excel's defaults).
export function outlinePlacement(state: LazyWorkbookState, sheetId: string): OutlinePlacement {
  const journal = state.editJournal.pageSetup.get(sheetId)
  const file = state.sheetFilePageSetups.get(sheetId)
  return {
    summaryBelow: journal?.outlineSummaryBelow ?? file?.outlineSummaryBelow ?? true,
    summaryRight: journal?.outlineSummaryRight ?? file?.outlineSummaryRight ?? true,
  }
}

/// Outline Settings: journals summary-below/right into the sheet's page
/// setup (the save writes `<sheetPr><outlinePr …/>`); returns null on
/// success or an error message.
export function handleOutlineSettings(
  ctx: DataToolsContext,
  placement: OutlinePlacement,
): string | null {
  const state = ctx.lazyWorkbookRef.current
  if (!state) return t('appOutlineNeedsFile')
  const sheetId = ctx.univerRef.current?.univerAPI
    .getActiveWorkbook()
    ?.getActiveSheet()
    ?.getSheetId()
  if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) {
    return t('appActiveSheetUnavailable')
  }
  const previous = outlinePlacement(state, sheetId)
  recordPageSetup(state.editJournal, sheetId, {
    outlineSummaryBelow: placement.summaryBelow,
    outlineSummaryRight: placement.summaryRight,
  })
  const undoPlacement = previous
  const redoPlacement = placement
  const sheetIdForUndo = sheetId
  const journal = state.editJournal
  const runtime = ctx.univerRef.current
  if (runtime) {
    pushVisualUndo(runtime, {
      undo: () => {
        recordPageSetup(journal, sheetIdForUndo, {
          outlineSummaryBelow: undoPlacement.summaryBelow,
          outlineSummaryRight: undoPlacement.summaryRight,
        })
        ctx.setPendingEdits(journalSize(journal))
        ctx.onOutlineChanged?.()
      },
      redo: () => {
        recordPageSetup(journal, sheetIdForUndo, {
          outlineSummaryBelow: redoPlacement.summaryBelow,
          outlineSummaryRight: redoPlacement.summaryRight,
        })
        ctx.setPendingEdits(journalSize(journal))
        ctx.onOutlineChanged?.()
      },
    })
  }
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.onOutlineChanged?.()
  ctx.setMessage(t('appOutlineSettingsSaved'))
  return null
}

/// Hides or shows one outline group's detail span (the gutter's +/- click):
/// the span goes through the normal hidden pipeline (journaled by the
/// mutation listener, undoable through the combined visual-undo step below)
/// and the summary line's collapsed flag journals declaratively.
export function toggleOutlineGroup(
  ctx: DataToolsContext,
  axis: 'rows' | 'cols',
  detail: { start: number; end: number },
  summary: number,
  collapse: boolean,
): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !state || !worksheet) return
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const outline = sheetOutline(state, sheetId)
  const entries = axis === 'rows' ? outline.rows : outline.cols
  const kind = axis === 'rows' ? ('set-rows-outline' as const) : ('set-cols-outline' as const)
  const count = detail.end - detail.start + 1
  const summaryLevel = entries.get(summary)?.level ?? 0
  const wasCollapsed = entries.get(summary)?.collapsed ?? false

  const run = (hide: boolean): void => {
    // The hide/show commands journal through the App's mutation listener;
    // their native undo entries are dropped (outlineUndoGate) so the whole
    // collapse reverts as one visual-undo step.
    outlineUndoGate.active = true
    try {
      if (axis === 'rows') {
        if (hide) worksheet.hideRows(detail.start, count)
        else worksheet.showRows(detail.start, count)
      } else if (hide) {
        worksheet.hideColumns(detail.start, count)
      } else {
        worksheet.showColumns(detail.start, count)
      }
    } finally {
      outlineUndoGate.active = false
    }
    entries.set(summary, { level: summaryLevel, collapsed: hide })
    recordStructuralOp(state.editJournal, sheetId, {
      kind,
      start: summary,
      end: summary,
      level: summaryLevel,
      collapsed: hide,
    })
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.onOutlineChanged?.()
  }

  run(collapse)
  pushVisualUndo(runtime, {
    // Undo restores the group's previous collapsed state in full — the
    // hidden span and the summary flag together.
    undo: () => run(wasCollapsed),
    redo: () => run(collapse),
  })
  ctx.setMessage(collapse ? t('appDetailHidden') : t('appDetailShown'))
}

/// Outline groups: level edits journal declaratively (Univer has no outline
/// model) and land on the visual undo stack; Hide/Show Detail rides the
/// hidden pipeline plus a collapsed flag on the summary line (above or
/// below per the sheet's outlinePr).
export function handleOutline(
  ctx: DataToolsContext,
  action: 'group' | 'ungroup' | 'hide-detail' | 'show-detail',
  axis: 'rows' | 'cols',
): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appOutlineNeedsFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appOutlineSelectFirst'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const start = axis === 'rows' ? range.getRow() : range.getColumn()
  const end = axis === 'rows' ? start + range.getHeight() - 1 : start + range.getWidth() - 1
  const outline = sheetOutline(state, sheetId)
  const entries = axis === 'rows' ? outline.rows : outline.cols
  const kind = axis === 'rows' ? ('set-rows-outline' as const) : ('set-cols-outline' as const)

  if (action === 'hide-detail' || action === 'show-detail') {
    // The selection is the group's detail span; the summary line sits after
    // it (summary below/right) or before it, per the sheet's outlinePr.
    const summaryAfter = placementForAxis(outlinePlacement(state, sheetId), axis)
    const summary = summaryAfter ? end + 1 : start - 1
    if (summary < 0) {
      ctx.setMessage(t('appOutlineNoSummaryLine'))
      return
    }
    toggleOutlineGroup(ctx, axis, { start, end }, summary, action === 'hide-detail')
    return
  }

  // Group/Ungroup shifts each contiguous run of equal levels by ±1
  // (levels clamp to 0-7). Runs already at the boundary are skipped.
  const delta = action === 'group' ? 1 : -1
  const ops: { start: number; end: number; level: number; previous: number; collapsed: boolean }[] =
    []
  let runStart = start
  let runLevel = entries.get(start)?.level ?? 0
  let runCollapsed = entries.get(start)?.collapsed ?? false
  const closeRun = (runEnd: number): void => {
    const level = Math.min(MAX_OUTLINE_LEVEL, Math.max(0, runLevel + delta))
    if (level !== runLevel)
      ops.push({ start: runStart, end: runEnd, level, previous: runLevel, collapsed: runCollapsed })
  }
  for (let index = start + 1; index <= end; index += 1) {
    const level = entries.get(index)?.level ?? 0
    if (level !== runLevel) {
      closeRun(index - 1)
      runStart = index
      runLevel = level
      runCollapsed = entries.get(index)?.collapsed ?? false
    }
  }
  closeRun(end)
  if (ops.length === 0) {
    ctx.setMessage(action === 'group' ? t('appOutlineMaxLevel') : t('appNothingToUngroup'))
    return
  }
  const applyLevels = (direction: 1 | -1): void => {
    for (const op of ops) {
      const level = direction === 1 ? op.level : op.previous
      for (let index = op.start; index <= op.end; index += 1) {
        entries.set(index, {
          level,
          collapsed: entries.get(index)?.collapsed ?? false,
        })
      }
      recordStructuralOp(state.editJournal, sheetId, {
        kind,
        start: op.start,
        end: op.end,
        level,
      })
    }
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.onOutlineChanged?.()
  }
  applyLevels(1)
  pushVisualUndo(runtime, {
    undo: () => applyLevels(-1),
    redo: () => applyLevels(1),
  })
  ctx.setMessage(
    action === 'group'
      ? axis === 'rows'
        ? t('appRowsGrouped')
        : t('appColsGrouped')
      : axis === 'rows'
        ? t('appRowsUngrouped')
        : t('appColsUngrouped'),
  )
}

/// Text to Columns: the raw material for the wizard's preview — the single
/// selected column's text, plus where it sits. null when the selection is
/// not exactly one column (the caller explains instead of opening).
export function readTextToColumnsSource(ctx: DataToolsContext): {
  rows: readonly string[]
  startRow: number
  startColumn: number
  destinationLabel: string
} | null {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range || range.getWidth() !== 1) return null
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const text = worksheet
    .getRange(startRow, startColumn, range.getHeight(), 1)
    .getDisplayValues()
    .map((row) => String(row[0] ?? ''))
  return {
    rows: text,
    startRow,
    startColumn,
    destinationLabel: `${columnLetter(startColumn)}${startRow + 1}`,
  }
}

/// The wizard's Finish: splits the selected column per the parsed config
/// and lands the fields at the destination through the normal journaled
/// write channel. Date-typed columns additionally carry a date number
/// format. Returns null on success or a user-facing error message.
export function handleTextToColumns(
  ctx: DataToolsContext,
  config: TextToColumnsConfig,
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) return t('appSelectCellFirst')
  if (range.getWidth() !== 1) return t('appTextToColsSelectOne')
  const source = readTextToColumnsSource(ctx)
  if (!source) return t('appTextToColsSelectOne')
  if (source.rows.length === 0) return t('appTextToColsEmpty')

  const delimiters = activeDelimiterChars(config.delimiters)
  if (config.mode === 'delimited' && delimiters.length === 0) {
    return t('appTextToColsNeedDelimiter')
  }
  const split = (text: string): string[] =>
    config.mode === 'fixed-width'
      ? splitFixedWidth(text, config.breaks)
      : splitDelimited(text, delimiters, config.delimiters.consecutiveAsOne)

  const fields = source.rows.map(split)
  const width = fields.reduce((max, row) => Math.max(max, row.length), 0)
  if (width === 0) return t('appTextToColsEmpty')
  const destination = config.destination
    ? parseDestinationCell(config.destination)
    : { row: source.startRow, column: source.startColumn }
  if (!destination) return t('appTextToColsBadDestination', { ref: config.destination ?? '' })

  const values = fields.map((row) =>
    Array.from({ length: width }, (_, column) => {
      const type = config.columnTypes[column] ?? 'general'
      return coerceFieldValue(row[column] ?? '', type)
    }),
  )
  try {
    worksheet.getRange(destination.row, destination.column, fields.length, width).setValues(values)
    for (let column = 0; column < width; column += 1) {
      if (DATE_COLUMN_TYPES.includes(config.columnTypes[column] ?? 'general')) {
        worksheet
          .getRange(destination.row, destination.column + column, fields.length, 1)
          .setNumberFormat('yyyy-mm-dd')
      }
    }
  } catch (error: unknown) {
    return error instanceof Error ? error.message : t('appCommandFailed')
  }
  ctx.setMessage(
    t('appTextToColsDone', {
      rows: fields.length,
      columns: width,
      cell: `${columnLetter(destination.column)}${destination.row + 1}`,
    }),
  )
  return null
}

/// Home → Format as Table: the manual entry over the same engine as the
/// AI add_table op (Univer table rendering + journal + native table part).
export function handleFormatAsTable(ctx: DataToolsContext, style: string): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appTablesNeedFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appTableSelectRange'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const endRow = startRow + range.getHeight() - 1
  const endColumn = startColumn + range.getWidth() - 1
  try {
    applyAiTableAdd(runtime, state, {
      op: 'add_table',
      sheetId,
      range: `${columnLabel(startColumn)}${startRow + 1}:${columnLabel(endColumn)}${endRow + 1}`,
      style,
      bandedRows: true,
    })
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appTableCreated'))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appTableCreateFailed'))
  }
}

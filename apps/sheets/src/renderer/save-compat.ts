/**
 * Registry of workbook constructs whose save is fail-closed (PAR-206).
 *
 * A fail-closed construct is a feature the user can build on screen that the
 * save pipeline cannot write back: either every save is refused until the
 * construct is undone/deleted, or the save would silently drop it. The save
 * refuses with a clear error only AFTER the user invested the work (and the
 * error surfaces in the status bar at the ⌘S moment). This module is the
 * single source of truth both for the pre-flight banner (which constructs are
 * currently in the workbook) and, in the future, for the save pipeline's own
 * feature checks — one inventory, one wording, one place to extend.
 *
 * Inventory of the fail-closed save scenarios (gateway refusal sites):
 * - Univer-only multi-select list data validation: every save fails until
 *   the rule is deleted (gateway/xlsx-dv.ts:74, error key appSaveErrMultiSelectList).
 * - Edits to extended (x14) conditional formatting: the save refuses when a
 *   dirty sheet's x14 base half drifted or a linked rule was removed
 *   (gateway/xlsx-cf.ts:163, appSaveErrX14Cf).
 * - Edits to extended (x14) data validation: the save refuses whenever a
 *   dirty sheet carries any x14 rule (gateway/xlsx-dv.ts:42, appSaveErrX14Dv).
 * - Structural work entangled with session additions on session-added
 *   sheets: the two-phase save strands the held addition and refuses
 *   (renderer/save-actions.ts stranded check, gateway/xlsx-gateway.ts:643/652/664,
 *   appSaveHeldStranded).
 * - Row/column moves and range moves that tear a table: header/totals row
 *   relocation, torn swap spans, merge over a table, range-move contact
 *   (gateway/xlsx-structure.ts:729/1008/1013/1029/1032/1052, appSaveErrMoveOverlap).
 * - A range move over the sheet's auto-filter pins coordinates the move
 *   would invalidate (gateway/xlsx-structure.ts:1803).
 * - Duplicating a sheet that carries drawings, tables, or pivots: the clone
 *   cannot share those parts and the save refuses
 *   (gateway/xlsx-sheets.ts:195, docs/compatibility.md "Sheet duplication").
 * - Chart edits outside the supported envelope refuse the save but are
 *   already blocked at the edit gate (ChartPanels `supported`), so no
 *   renderer detector is needed here (gateway/xlsx-chart.ts:5).
 * - Oversized anchored parts / missing relationships refuse row/column
 *   shifts (gateway/xlsx-gateway.ts:1764) — environmental, not a user-built
 *   construct; listed for completeness.
 * - CSV sessions: saving keeps plain values of the active sheet only —
 *   formulas, formatting, and other sheets do not survive the save
 *   (main/sheets-main.ts:306; the per-save confirm dialog asks at ⌘S, the
 *   banner shows it up front).
 *
 * Detectors run on renderer-visible state only (edit journal + file
 * metadata + the live Univer model). They may produce false negatives (a
 * construct the renderer cannot see yet) but must never produce false
 * positives: the banner is a pre-flight warning, never a blocker, and a
 * wrong alarm would erode trust in it.
 */
import {
  isSheetRemoved,
  toSavePivotAdds,
  toSaveSheetOps,
  toSaveStructuralOps,
  toSaveTableAdds,
  type StructuralJournalOp,
} from './edit-journal'
import type { StringKey } from './i18n/locale'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

export type SaveCompatFeatureId =
  | 'multi-select-dv'
  | 'stranded-addition'
  | 'structure-conflict'
  | 'duplicate-carries-parts'
  | 'x14-cf-edit'
  | 'x14-dv-edit'
  | 'csv-flatten'

/// What one detector found: the feature and, when known, the sheet or item
/// it came from (surfaced as "{item}" detail next to the feature label).
export interface SaveCompatFinding {
  readonly id: SaveCompatFeatureId
  readonly detail?: string
}

/// Everything a detector may look at; built fresh per banner evaluation.
export interface SaveCompatContext {
  readonly state: LazyWorkbookState
  readonly runtime: UniverRuntime | null
}

export interface SaveCompatFeature {
  readonly id: SaveCompatFeatureId
  /// i18n key (app domain) of the user-facing description of the construct.
  readonly labelKey: StringKey
  /// The fail-closed refusal sites this feature mirrors (inventory anchor
  /// for reviewers; see the module doc comment for the full map).
  readonly refs: readonly string[]
  /// Renderer-side detector, or null when the renderer cannot observe the
  /// construct yet (the future save-side check lives in the gateway, which
  /// sees the raw package parts). Null detectors never fire.
  readonly detect: ((ctx: SaveCompatContext) => SaveCompatFinding | null) | null
}

interface Span {
  start: number
  end: number
}

interface FileTableSpan {
  name: string
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  headerRows: number
  totalsRows: number
}

function spansOverlap(aStart: number, aEnd: number, bStart: number, bEnd: number): boolean {
  return aStart <= bEnd && bStart <= aEnd
}

function areasOverlap(
  a: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  b: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  return (
    spansOverlap(a.startRow, a.endRow, b.startRow, b.endRow) &&
    spansOverlap(a.startColumn, a.endColumn, b.startColumn, b.endColumn)
  )
}

function sheetName(state: LazyWorkbookState, sheetId: string): string | undefined {
  return state.file.sheets.find((sheet) => sheet.id === sheetId)?.name
}

/// Univer-only multi-select list validation rules cannot be serialized to
/// xlsx; the save fails until the rule is deleted. The sheet carrying one is
/// necessarily dvDirty (the rule was created this session), so scanning the
/// dirty sheets' live rules is exact.
function detectMultiSelectDv({ state, runtime }: SaveCompatContext): SaveCompatFinding | null {
  const dirty = state.editJournal.dvDirty
  if (dirty.size === 0 || !runtime) return null
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return null
  for (const sheetId of dirty) {
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const worksheet = workbook.getSheetBySheetId(sheetId)
    if (!worksheet) continue
    try {
      const rules = (
        worksheet as unknown as {
          getDataValidations(): { rule: { type?: unknown } }[]
        }
      ).getDataValidations()
      if (rules.some(({ rule }) => rule?.type === 'listMultiple')) {
        const detail = sheetName(state, sheetId)
        return detail === undefined ? { id: 'multi-select-dv' } : { id: 'multi-select-dv', detail }
      }
    } catch {
      // The model can be mid-mutation during a render; a failed read must
      // never nag the user, so treat it as "nothing seen".
      return null
    }
  }
  return null
}

/// The save holds new pivots/tables back when structural or sheet changes
/// ride in the same pass (their coordinates entangle) and saves them in a
/// second phase — except additions on session-added sheets, which the
/// reopened session cannot address. Mirror the save flow's stranded check
/// exactly (save-actions.ts), or the banner would disagree with ⌘S.
function detectStrandedAddition({ state }: SaveCompatContext): SaveCompatFinding | null {
  const journal = state.editJournal
  const structuralOps = toSaveStructuralOps(journal)
  if (structuralOps.length === 0 && toSaveSheetOps(journal).length === 0) return null
  const heldPivots = toSavePivotAdds(journal)
  const heldTables = structuralOps.length > 0 ? toSaveTableAdds(journal) : []
  const added = journal.sheets.added
  const stranded =
    heldPivots.some((pivot) => added.has(pivot.sheetId) || added.has(pivot.sourceSheetId)) ||
    heldTables.some((table) => added.has(table.sheetId))
  return stranded ? { id: 'stranded-addition' } : null
}

/// The two adjacent pre-move blocks of a whole-row/column move, mirroring the
/// gateway's BlockSwap (xlsx-structure.ts) and view-transform.ts swapBlocks.
function swapBlocks(op: Extract<StructuralJournalOp, { before: number }>): {
  first: Span
  second: Span
} {
  const first =
    op.before > op.index
      ? { start: op.index, end: op.index + op.count - 1 }
      : { start: op.before, end: op.index - 1 }
  const second =
    op.before > op.index
      ? { start: op.index + op.count, end: op.before - 1 }
      : { start: op.index, end: op.index + op.count - 1 }
  return { first, second }
}

/// Gateway moveRange for a linear insert/delete shift on one axis
/// (xlsx-structure.ts): null when the whole span was deleted.
function shiftSpan(span: Span, op: { index: number; count: number }, insert: boolean): Span | null {
  if (insert) {
    return {
      start: span.start >= op.index ? span.start + op.count : span.start,
      end: span.end >= op.index ? span.end + op.count : span.end,
    }
  }
  const deleted = { start: op.index, end: op.index + op.count - 1 }
  if (span.start >= deleted.start && span.end <= deleted.end) return null
  const delta = -op.count
  const start =
    span.start > deleted.end
      ? span.start + delta
      : span.start >= deleted.start
        ? deleted.start
        : span.start
  const end =
    span.end > deleted.end
      ? span.end + delta
      : span.end >= deleted.start
        ? deleted.start - 1
        : span.end
  return end < start ? null : { start, end }
}

/// Gateway moveRange three-way rule for swap shifts (xlsx-structure.ts): a
/// span fully inside one block relocates, a span containing or missing the
/// envelope stays put, and anything partially overlapping the blocks is torn
/// — the XML has no faithful single image of it.
function swapSpan(span: Span, blocks: { first: Span; second: Span }): Span | null {
  const { first, second } = blocks
  if (span.end < first.start || span.start > second.end) return span
  if (span.start <= first.start && span.end >= second.end) return span
  if (span.start >= first.start && span.end <= first.end) {
    const delta = second.end - second.start + 1
    return { start: span.start + delta, end: span.end + delta }
  }
  if (span.start >= second.start && span.end <= second.end) {
    const delta = first.end - first.start + 1
    return { start: span.start - delta, end: span.end - delta }
  }
  return null
}

/// Replays one sheet's journaled structural ops over the file-side table
/// spans exactly the way shiftTablePart does at save time, returning the
/// first construct the save would refuse on. Op coordinates are shared with
/// the gateway's sequential replay (view-transform.ts maps the same stream),
/// so mirroring the span math keeps the banner and ⌘S in step.
function findTableConflict(
  ops: readonly StructuralJournalOp[],
  tables: FileTableSpan[],
  worksheetFilter: {
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
  } | null,
): string | null {
  for (const op of ops) {
    for (const table of tables) {
      if (op.kind === 'merge-cells' && areasOverlap(op.range, table)) {
        return table.name
      }
      if (op.kind === 'move-range') {
        // Any contact with either rectangle refuses (v1 fail-closed family,
        // xlsx-structure.ts "A range move overlaps table").
        if (areasOverlap(op.from, table) || areasOverlap(op.to, table)) return table.name
        continue
      }
      if (!('index' in op)) continue
      if (op.kind === 'insert-rows' || op.kind === 'remove-rows') {
        const refused = rowShiftRefusal(table, op, op.kind === 'insert-rows')
        if (refused) return refused
      } else if (op.kind === 'move-rows') {
        const refused = rowSwapRefusal(table, swapBlocks(op))
        if (refused) return refused
      } else if (op.kind === 'move-cols') {
        const refused = columnSwapRefusal(table, swapBlocks(op))
        if (refused) return refused
      }
    }
    // The sheet's own auto-filter pins its coordinates against range moves
    // (xlsx-structure.ts "A range move overlaps the sheet's auto-filter").
    // Its live range is screen-space; exact only while no shifting op
    // preceded the move — a false negative afterwards, never a false alarm.
    if (worksheetFilter && op.kind === 'move-range') {
      if (areasOverlap(op.from, worksheetFilter) || areasOverlap(op.to, worksheetFilter)) {
        return 'autoFilter'
      }
    }
    // Advance the replayed spans through this op (a replace move shifts no
    // axis line — the view-transform invariant — so move-range is a no-op).
    if (!('index' in op)) continue
    if (op.kind === 'insert-rows' || op.kind === 'remove-rows') {
      advanceSpans(tables, op, 'row', op.kind === 'insert-rows')
    } else if (op.kind === 'insert-cols' || op.kind === 'remove-cols') {
      advanceSpans(tables, op, 'column', op.kind === 'insert-cols')
    } else if (op.kind === 'move-rows' || op.kind === 'move-cols') {
      const blocks = swapBlocks(op)
      const axisRows = op.kind === 'move-rows'
      for (const table of tables) {
        const span = axisRows
          ? swapSpan({ start: table.startRow, end: table.endRow }, blocks)
          : swapSpan({ start: table.startColumn, end: table.endColumn }, blocks)
        if (span === null) continue
        if (axisRows) {
          table.startRow = span.start
          table.endRow = span.end
        } else {
          table.startColumn = span.start
          table.endColumn = span.end
        }
      }
    }
  }
  return null
}

/// assertTableRowShiftSupported (xlsx-structure.ts): deleting a table's
/// header or totals rows, or leaving it without data rows, refuses the save.
function rowShiftRefusal(
  table: FileTableSpan,
  op: { index: number; count: number },
  insert: boolean,
): string | null {
  if (insert) return null
  const deleted = { start: op.index, end: op.index + op.count - 1 }
  if (
    table.headerRows > 0 &&
    spansOverlap(deleted.start, deleted.end, table.startRow, table.startRow + table.headerRows - 1)
  ) {
    return table.name
  }
  if (
    table.totalsRows > 0 &&
    spansOverlap(deleted.start, deleted.end, table.endRow - table.totalsRows + 1, table.endRow)
  ) {
    return table.name
  }
  const moved = shiftSpan({ start: table.startRow, end: table.endRow }, op, false)
  const dataRows =
    moved === null ? 0 : moved.end - moved.start + 1 - table.headerRows - table.totalsRows
  return dataRows < 1 ? table.name : null
}

/// assertTableRowMoveSupported + moveRange tear for whole-row swaps: a table
/// fully inside one block relocates and a table containing the whole swap
/// stays put (an interior data reorder); relocating its header or totals row
/// relative to the data refuses, as does any boundary tear the ref move
/// cannot map (xlsx-structure.ts assertTableRowMoveSupported + moveRange).
function rowSwapRefusal(
  table: FileTableSpan,
  blocks: { first: Span; second: Span },
): string | null {
  const within = (span: Span): boolean => table.startRow >= span.start && table.endRow <= span.end
  if (within(blocks.first) || within(blocks.second)) return null
  if (!spansOverlap(table.startRow, table.endRow, blocks.first.start, blocks.second.end)) {
    return null
  }
  const envelopeStart = blocks.first.start
  const envelopeEnd = blocks.second.end
  if (
    table.headerRows > 0 &&
    spansOverlap(envelopeStart, envelopeEnd, table.startRow, table.startRow + table.headerRows - 1)
  ) {
    return table.name
  }
  if (
    table.totalsRows > 0 &&
    spansOverlap(envelopeStart, envelopeEnd, table.endRow - table.totalsRows + 1, table.endRow)
  ) {
    return table.name
  }
  return swapSpan({ start: table.startRow, end: table.endRow }, blocks) === null ? table.name : null
}

/// assertTableColumnMoveSupported: a table whose column span sits fully
/// inside one swapped block relocates wholesale; any other envelope overlap
/// would reorder its interior columns and desync the tableColumn list.
function columnSwapRefusal(
  table: FileTableSpan,
  blocks: { first: Span; second: Span },
): string | null {
  const within = (span: Span): boolean =>
    table.startColumn >= span.start && table.endColumn <= span.end
  if (within(blocks.first) || within(blocks.second)) return null
  if (!spansOverlap(table.startColumn, table.endColumn, blocks.first.start, blocks.second.end)) {
    return null
  }
  return table.name
}

function advanceSpans(
  tables: FileTableSpan[],
  op: { index: number; count: number },
  axis: 'row' | 'column',
  insert: boolean,
): void {
  for (const table of tables) {
    const span =
      axis === 'row'
        ? shiftSpan({ start: table.startRow, end: table.endRow }, op, insert)
        : shiftSpan({ start: table.startColumn, end: table.endColumn }, op, insert)
    if (span === null) continue
    if (axis === 'row') {
      table.startRow = span.start
      table.endRow = span.end
    } else {
      table.startColumn = span.start
      table.endColumn = span.end
    }
  }
}

/// Row/column moves, range moves, and merges that tear a file-side table, or
/// a range move that overlaps the sheet's auto-filter: the save refuses.
function detectStructureConflict({ state }: SaveCompatContext): SaveCompatFinding | null {
  for (const [sheetId, ops] of state.editJournal.structuralOps) {
    if (ops.length === 0 || isSheetRemoved(state.editJournal, sheetId)) continue
    const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
    // Session-added sheets have no file parts to tear (their tables are a
    // different feature — the stranded-addition check).
    if (!sheet) continue
    const tables: FileTableSpan[] = sheet.tables.map((table) => ({
      name: table.name ?? 'table',
      startRow: table.range.startRow,
      endRow: table.range.endRow,
      startColumn: table.range.startColumn,
      endColumn: table.range.endColumn,
      headerRows: table.headerRowCount,
      totalsRows: table.totalsRowCount ?? 0,
    }))
    const filterOrigin = state.filterOrigins.get(sheetId)
    const worksheetFilter =
      filterOrigin !== undefined && filterOrigin.origin === 'worksheet'
        ? {
            startRow: filterOrigin.range.startRow,
            endRow: filterOrigin.range.endRow,
            startColumn: filterOrigin.range.startColumn,
            endColumn: filterOrigin.range.endColumn,
          }
        : null
    const conflicted = findTableConflict(ops, tables, worksheetFilter)
    if (conflicted !== null) {
      return { id: 'structure-conflict', detail: sheet.name }
    }
  }
  return null
}

/// Duplicating a sheet that carries charts/images/shapes, tables, pivots, or
/// sheet-scoped defined names refuses at save: the clone cannot share those
/// parts (xlsx-sheets.ts prepareClonedSheetRels).
function detectDuplicateCarriesParts({ state }: SaveCompatContext): SaveCompatFinding | null {
  const journal = state.editJournal
  for (const [sheetId, added] of journal.sheets.added) {
    if (added.sourceSheetId === undefined || isSheetRemoved(journal, sheetId)) continue
    const sourceIndex = state.file.sheets.findIndex((sheet) => sheet.id === added.sourceSheetId)
    const source = sourceIndex === -1 ? undefined : state.file.sheets[sourceIndex]
    if (source === undefined) continue
    const carriesParts =
      source.tables.length > 0 ||
      source.pivotRanges.length > 0 ||
      state.file.visuals.some((visual) => visual.sheetId === source.id) ||
      state.file.definedNames.some((name) => name.sheetIndex === sourceIndex)
    if (carriesParts) {
      return { id: 'duplicate-carries-parts', detail: added.name }
    }
  }
  return null
}

function detectCsvFlatten({ state }: SaveCompatContext): SaveCompatFinding | null {
  return state.file.csvPath !== undefined ? { id: 'csv-flatten' } : null
}

/// The registry: every known fail-closed construct, detectable or not. The
/// banner renders findings in this order; the future save pipeline reads the
/// same entries for its own feature checks.
export const SAVE_COMPAT_FEATURES: readonly SaveCompatFeature[] = [
  {
    id: 'csv-flatten',
    labelKey: 'appSaveCompatCsv',
    refs: [
      'apps/sheets/src/main/sheets-main.ts:306',
      'apps/sheets/src/renderer/save-actions.ts:252',
    ],
    detect: detectCsvFlatten,
  },
  {
    id: 'multi-select-dv',
    labelKey: 'appSaveCompatMultiSelectDv',
    refs: ['apps/sheets/src/gateway/xlsx-dv.ts:74'],
    detect: detectMultiSelectDv,
  },
  {
    id: 'stranded-addition',
    labelKey: 'appSaveCompatStranded',
    refs: [
      'apps/sheets/src/renderer/save-actions.ts:204',
      'apps/sheets/src/gateway/xlsx-gateway.ts:643',
      'apps/sheets/src/gateway/xlsx-gateway.ts:664',
    ],
    detect: detectStrandedAddition,
  },
  {
    id: 'structure-conflict',
    labelKey: 'appSaveCompatStructureConflict',
    refs: [
      'apps/sheets/src/gateway/xlsx-structure.ts:729',
      'apps/sheets/src/gateway/xlsx-structure.ts:1008',
      'apps/sheets/src/gateway/xlsx-structure.ts:1032',
      'apps/sheets/src/gateway/xlsx-structure.ts:1052',
      'apps/sheets/src/gateway/xlsx-structure.ts:1803',
    ],
    detect: detectStructureConflict,
  },
  {
    id: 'duplicate-carries-parts',
    labelKey: 'appSaveCompatDuplicateParts',
    refs: ['apps/sheets/src/gateway/xlsx-sheets.ts:195'],
    detect: detectDuplicateCarriesParts,
  },
  {
    // The renderer cannot see a sheet's x14 conditional formatting (it lives
    // in the worksheet extLst, read sidecar-side); the save-side check is the
    // registry consumer until a file-side marker exists.
    id: 'x14-cf-edit',
    labelKey: 'appSaveCompatX14Cf',
    refs: ['apps/sheets/src/gateway/xlsx-cf.ts:163'],
    detect: null,
  },
  {
    id: 'x14-dv-edit',
    labelKey: 'appSaveCompatX14Dv',
    refs: ['apps/sheets/src/gateway/xlsx-dv.ts:42'],
    detect: null,
  },
]

/// Runs every detector and collects the findings, in registry order. A
/// detector failure (unexpected model shape) must never break the banner:
/// it degrades to "nothing seen".
export function collectSaveCompatFindings(ctx: SaveCompatContext): SaveCompatFinding[] {
  const findings: SaveCompatFinding[] = []
  for (const feature of SAVE_COMPAT_FEATURES) {
    if (feature.detect === null) continue
    try {
      const finding = feature.detect(ctx)
      if (finding !== null) findings.push(finding)
    } catch {
      // Degrade silently — see the module doc comment on false positives.
    }
  }
  return findings
}

/// Stable identity of the currently visible finding set: dismissing hides
/// exactly this set for the session; a different set appearing later re-shows
/// the banner (a new construct arrived).
export function saveCompatDismissKey(findings: readonly SaveCompatFinding[]): string {
  return findings
    .map((finding) => finding.id)
    .sort()
    .join(',')
}

/// The banner's appear/disappear/dismiss policy: findings are visible unless
/// the user dismissed exactly this set for the session. An empty finding set
/// (construct resolved) is always invisible; a different set re-shows.
export function saveCompatVisible(
  findings: readonly SaveCompatFinding[],
  dismissedKey: string | null,
): readonly SaveCompatFinding[] {
  return saveCompatDismissKey(findings) === dismissedKey ? [] : findings
}

/// The i18n label of a feature, read from the registry — the banner and any
/// future save-side messaging share one wording per construct.
export function saveCompatLabelKey(id: SaveCompatFeatureId): StringKey {
  const feature = SAVE_COMPAT_FEATURES.find((candidate) => candidate.id === id)
  if (!feature) throw new Error(`Unknown save-compat feature: ${id}`)
  return feature.labelKey
}

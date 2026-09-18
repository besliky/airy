/// Allow-edit ranges are kept in current screen coordinates on
/// LazyWorkbookState: mapped from file space once when a sheet finishes
/// indexing, then incrementally through each recorded structural op — so the
/// dialog and the save read them without further translation (the save's
/// protectedRanges rewrite runs after structural replay).

import { columnLabel, parseRange } from '../domain/cell-address'
import type { StructuralJournalOp } from './edit-journal'
import { fileRangeToScreenRange, fileSpanToScreenEnvelope, fileToScreen } from './view-transform'

export interface ProtectedRangeEntry {
  readonly name: string
  readonly sqref: string
  readonly hasPassword: boolean
}

function formatArea(area: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}): string {
  const start = `${columnLabel(area.startColumn)}${area.startRow + 1}`
  if (area.startRow === area.endRow && area.startColumn === area.endColumn) return start
  return `${start}:${columnLabel(area.endColumn)}${area.endRow + 1}`
}

interface Area {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

/// The range envelope over-reads on a move that partially overlaps the area
/// (its span semantics serve viewport fetches) — for an edit whitelist that
/// would fail open. With a move in play on an axis, that axis's lines are
/// mapped one by one and the survivors split into contiguous runs instead;
/// the other axis keeps its envelope.
const EXACT_MOVE_LINE_CAP = 50_000

interface LineSpan {
  start: number
  end: number
}

/// Exact screen images of [start, end] as contiguous runs (empty when every
/// line was deleted). Screen positions are sorted before splitting: a move's
/// images are not monotonic in the file order.
function survivorRuns(
  ops: readonly StructuralJournalOp[],
  axis: 'row' | 'column',
  start: number,
  end: number,
): LineSpan[] {
  const screens: number[] = []
  for (let line = start; line <= end; line += 1) {
    const screen = fileToScreen(ops, axis, line)
    if (screen !== null) screens.push(screen)
  }
  screens.sort((a, b) => a - b)
  const runs: LineSpan[] = []
  for (const screen of screens) {
    const last = runs[runs.length - 1]
    if (last !== undefined && screen === last.end + 1) last.end = screen
    else runs.push({ start: screen, end: screen })
  }
  return runs
}

function mapArea(area: Area, ops: readonly StructuralJournalOp[]): Area[] {
  const exactRows =
    ops.some((op) => op.kind === 'move-rows') && area.endRow - area.startRow <= EXACT_MOVE_LINE_CAP
  const exactColumns =
    ops.some((op) => op.kind === 'move-cols') &&
    area.endColumn - area.startColumn <= EXACT_MOVE_LINE_CAP
  if (!exactRows && !exactColumns) {
    const moved = fileRangeToScreenRange(ops, area)
    return moved === null ? [] : [moved]
  }
  const rows = exactRows
    ? survivorRuns(ops, 'row', area.startRow, area.endRow)
    : envelopeRuns(ops, 'row', area.startRow, area.endRow)
  if (rows.length === 0) return []
  const columns = exactColumns
    ? survivorRuns(ops, 'column', area.startColumn, area.endColumn)
    : envelopeRuns(ops, 'column', area.startColumn, area.endColumn)
  if (columns.length === 0) return []
  // Both axes moved: the exact image is the cross product of their runs.
  const areas: Area[] = []
  for (const row of rows) {
    for (const column of columns) {
      areas.push({
        startRow: row.start,
        endRow: row.end,
        startColumn: column.start,
        endColumn: column.end,
      })
    }
  }
  return areas
}

/// The unmoved axis keeps the envelope: its lines were never shuffled, so
/// the bounding box of the survivors is the exact whitelisted span.
function envelopeRuns(
  ops: readonly StructuralJournalOp[],
  axis: 'row' | 'column',
  start: number,
  end: number,
): LineSpan[] {
  const envelope = fileSpanToScreenEnvelope(ops, axis, { start, end })
  return envelope === null ? [] : [envelope]
}

/// Maps a sqref (one or more space-separated A1 areas) through structural
/// ops; fully deleted areas drop out, an unparseable area stays verbatim.
/// null when nothing survives.
function mapSqref(sqref: string, ops: readonly StructuralJournalOp[]): string | null {
  const parts = sqref
    .split(/\s+/)
    .filter((part) => part !== '')
    .flatMap((part) => {
      let area
      try {
        area = parseRange(part.replaceAll('$', ''))
      } catch {
        return [part]
      }
      return mapArea(area, ops).map(formatArea)
    })
  return parts.length === 0 ? null : parts.join(' ')
}

/// Ranges with every area deleted drop out entirely.
export function mapProtectedRanges(
  ranges: readonly ProtectedRangeEntry[],
  ops: readonly StructuralJournalOp[],
): ProtectedRangeEntry[] {
  if (ops.length === 0) return [...ranges]
  const mapped: ProtectedRangeEntry[] = []
  for (const range of ranges) {
    const sqref = mapSqref(range.sqref, ops)
    if (sqref !== null) mapped.push({ ...range, sqref })
  }
  return mapped
}

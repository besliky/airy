// Single A1-notation parse/validate for the xlsx session, shared by the read
// and write paths (BUG-1632). The two sides used to diverge: journaling
// accepted whatever the regex parsed — "A0" journaled as row -1 and poisoned
// every later save ("Invalid cell coordinates: -1,0") — while reads refused
// the same ref with "Range is outside sheet". One parser plus one
// addressability predicate here means a ref the write path journals is
// exactly a ref the read path accepts, and vice versa.

/** SpreadsheetML hard grid limits: rows 1..1048576, columns A..XFD (16384). */
export const MAX_SHEET_ROWS = 1_048_576
export const MAX_SHEET_COLUMNS = 16_384

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/
const RANGE_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7}):\s*\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/

export function columnFromLabel(label: string): number {
  let value = 0
  for (const char of label.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64)
  return value - 1
}

export function columnToLabel(index: number): string {
  let label = ''
  let value = index + 1
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

/** An inclusive 0-based row/column window (single cell: start === end). */
export interface A1Range {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

/**
 * The one addressability predicate for parsed A1 coordinates: both corners
 * must land inside the SpreadsheetML grid. "A0" is valid A1 syntax but
 * addresses no cell (its 0-based row is -1), and refs past the grid's last
 * cell (row > ${MAX_SHEET_ROWS}, column past XFD) address nothing either.
 */
export function assertAddressableA1(range: A1Range, spec: string): void {
  if (range.startRow < 0 || range.startColumn < 0) {
    throw new Error(
      `Ref "${spec}" addresses no cell: rows and columns start at 1 in A1 notation ` +
        '(the first cell is "A1").',
    )
  }
  if (range.endRow >= MAX_SHEET_ROWS || range.endColumn >= MAX_SHEET_COLUMNS) {
    throw new Error(
      `Ref "${spec}" is past the sheet's last cell: rows run 1..${String(MAX_SHEET_ROWS)}, ` +
        'columns A..XFD.',
    )
  }
}

/** Parse "B2" or "A1:C10" into an inclusive 0-based range (unordered refs allowed). */
export function parseA1Range(spec: string): A1Range {
  const single = CELL_RE.exec(spec.trim())
  if (single) {
    const row = Number(single[2]) - 1
    const column = columnFromLabel(single[1]!)
    const range = { startRow: row, endRow: row, startColumn: column, endColumn: column }
    assertAddressableA1(range, spec)
    return range
  }
  const range = RANGE_RE.exec(spec.trim())
  if (!range) throw new Error(`Invalid A1-style range "${spec}" (expected e.g. "A1:C10" or "B2").`)
  const startColumn = columnFromLabel(range[1]!)
  const endColumn = columnFromLabel(range[3]!)
  const startRow = Number(range[2]) - 1
  const endRow = Number(range[4]) - 1
  const parsed = {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startColumn: Math.min(startColumn, endColumn),
    endColumn: Math.max(startColumn, endColumn),
  }
  assertAddressableA1(parsed, spec)
  return parsed
}

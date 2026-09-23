// Single A1-notation parse/validate for the xlsx session, shared by the read
// and write paths: both go through parseA1Range, so the accepted coordinate
// space cannot drift apart between journaling edits and reading them back.

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

/** Parse "B2" or "A1:C10" into an inclusive 0-based range (unordered refs allowed). */
export function parseA1Range(spec: string): A1Range {
  const single = CELL_RE.exec(spec.trim())
  if (single) {
    const row = Number(single[2]) - 1
    const column = columnFromLabel(single[1]!)
    return { startRow: row, endRow: row, startColumn: column, endColumn: column }
  }
  const range = RANGE_RE.exec(spec.trim())
  if (!range) throw new Error(`Invalid A1-style range "${spec}" (expected e.g. "A1:C10" or "B2").`)
  const startColumn = columnFromLabel(range[1]!)
  const endColumn = columnFromLabel(range[3]!)
  const startRow = Number(range[2]) - 1
  const endRow = Number(range[4]) - 1
  return {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startColumn: Math.min(startColumn, endColumn),
    endColumn: Math.max(startColumn, endColumn),
  }
}

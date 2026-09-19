/// Text to Columns: split one selected column by delimiters or fixed-width
/// break positions, coerce each field per its column type, and land the
/// result through the normal journaled write channel (cell values, plus a
/// number format for Date columns — both already ride the undo stack and
/// the streamed save journal).

/// Excel's wizard delimiters; `custom` is the "Other" box, where every
/// character splits.
export interface TextToColumnsDelimiters {
  readonly tab: boolean
  readonly semicolon: boolean
  readonly comma: boolean
  readonly space: boolean
  readonly custom: string
  /// Treat consecutive delimiters as one (skip the empty fields between).
  readonly consecutiveAsOne: boolean
}

export const DEFAULT_DELIMITERS: TextToColumnsDelimiters = {
  tab: false,
  semicolon: false,
  comma: true,
  space: false,
  custom: '',
  consecutiveAsOne: false,
}

export type TextToColumnsMode = 'delimited' | 'fixed-width'

/// Step-3 column formats. The date variants carry Excel's date-order
/// interpretation of the wizard (DMY / MDY / YMD).
export type TextToColumnType = 'general' | 'text' | 'date-dmy' | 'date-mdy' | 'date-ymd'

export const DATE_COLUMN_TYPES: readonly TextToColumnType[] = ['date-dmy', 'date-mdy', 'date-ymd']

export interface TextToColumnsConfig {
  readonly mode: TextToColumnsMode
  readonly delimiters: TextToColumnsDelimiters
  /// 1-based character positions after which a fixed-width break lands,
  /// ascending and unique.
  readonly breaks: readonly number[]
  /// One format per output column (missing entries default to General).
  readonly columnTypes: readonly TextToColumnType[]
  /// Destination top-left as an A1 reference on the active sheet; null
  /// writes back over the source column.
  readonly destination: string | null
}

/// The active delimiter characters, in Excel's checkbox order.
export function activeDelimiterChars(delimiters: TextToColumnsDelimiters): string[] {
  const chars: string[] = []
  if (delimiters.tab) chars.push('\t')
  if (delimiters.semicolon) chars.push(';')
  if (delimiters.comma) chars.push(',')
  if (delimiters.space) chars.push(' ')
  for (const char of delimiters.custom) {
    if (!chars.includes(char)) chars.push(char)
  }
  return chars
}

/// Splits one cell's text by any of the delimiter characters. With
/// consecutiveAsOne, runs of delimiters separate one field; without it,
/// adjacent delimiters keep their empty field between them (Excel's
/// behavior both ways).
export function splitDelimited(
  text: string,
  delimiters: readonly string[],
  consecutiveAsOne: boolean,
): string[] {
  if (delimiters.length === 0) return [text]
  const escaped = delimiters.map(escapeForRegex).join('')
  const pattern = consecutiveAsOne ? new RegExp(`[${escaped}]+`) : new RegExp(`[${escaped}]`)
  return text.split(pattern)
}

/// "5, 12 20" → [5, 12, 20]; null when any part is not a positive integer
/// (the dialog reports it). Sorted, deduped, capped like the wizard's ruler.
export function parseBreakPositions(input: string): number[] | null {
  const parts = input
    .split(/[\s,;]+/)
    .map((part) => part.trim())
    .filter((part) => part !== '')
  if (parts.length === 0) return []
  const positions: number[] = []
  for (const part of parts) {
    if (!/^\d{1,5}$/.test(part)) return null
    const position = Number(part)
    if (position < 1 || position > 32_767) return null
    positions.push(position)
  }
  const unique = [...new Set(positions)].sort((a, b) => a - b)
  return unique.slice(0, 100)
}

/// Splits one cell's text at the fixed-width break positions: the first
/// field covers characters [1..breaks[0]], the next (breaks[0]..breaks[1]],
/// and the tail runs to the end of the text.
export function splitFixedWidth(text: string, breaks: readonly number[]): string[] {
  if (breaks.length === 0) return [text]
  const fields: string[] = []
  let start = 0
  for (const position of breaks) {
    const end = Math.min(position, text.length)
    if (end < start) break
    fields.push(text.slice(start, end))
    start = end
    if (start >= text.length) break
  }
  if (start < text.length) fields.push(text.slice(start))
  return fields
}

/// One parsed cell ready to write: the value plus (for dates) the number
/// format applied to its column.
export interface ParsedFieldValue {
  readonly v: string | number | boolean | null
}

/// General typing follows Excel: numbers (with thousands separators),
/// percentages, TRUE/FALSE booleans; anything else stays text. Text keeps
/// the field verbatim; the date orders parse numeric d/m/y combinations
/// and non-dates stay text (Excel's result for an unparsable date, too).
export function coerceFieldValue(text: string, type: TextToColumnType): ParsedFieldValue {
  const trimmed = text.trim()
  if (trimmed === '') return { v: null }
  if (type === 'text') return { v: text }
  if (type === 'general') {
    const upper = trimmed.toUpperCase()
    if (upper === 'TRUE') return { v: true }
    if (upper === 'FALSE') return { v: false }
    const percent = /^([+-]?)\d+(\.\d+)?%$/.exec(trimmed)
    if (percent) {
      const fraction = Number(trimmed.replace(/%$/, '')) / 100
      if (Number.isFinite(fraction)) return { v: fraction }
    }
    if (/^[+-]?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(trimmed.replace(/,/g, ''))) {
      const numeric = Number(trimmed.replace(/,/g, ''))
      if (Number.isFinite(numeric)) return { v: numeric }
    }
    return { v: text }
  }
  const date = parseDateText(trimmed, type)
  return date === null ? { v: text } : { v: date }
}

/// Days since Excel's epoch. Dates from 1900-03-01 count from 1899-12-30;
/// before that they sit one lower, because Excel's fake 1900-02-29 (serial
/// 60) shifts everything after it — matching what Excel stores.
export function excelDateSerial(year: number, month: number, d: number): number {
  const epoch = Date.UTC(1899, 11, 30)
  const days = Math.round((Date.UTC(year, month - 1, d) - epoch) / 86_400_000)
  return days < 61 ? days - 1 : days
}

/// Numeric date forms with ., -, or / separators. Two-digit years follow
/// Excel: 00-29 → 2000-2029, 30-99 → 1930-1999.
function parseDateText(text: string, order: 'date-dmy' | 'date-mdy' | 'date-ymd'): number | null {
  const match = /^(\d{1,4})[.\-/](\d{1,2})[.\-/](\d{1,4})$/.exec(text)
  if (!match) return null
  const first = Number(match[1])
  const second = Number(match[2])
  const third = Number(match[3])
  let year: number
  let month: number
  let day: number
  if (order === 'date-ymd') {
    year = first
    month = second
    day = third
  } else if (order === 'date-mdy') {
    month = first
    day = second
    year = third
  } else {
    day = first
    month = second
    year = third
  }
  if (year < 100) year += year <= 29 ? 2000 : 1900
  if (month < 1 || month > 12 || day < 1 || day > 31) return null
  if (year < 1900 || year > 9999) return null
  const daysInMonth = [31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  if (day > (daysInMonth[month - 1] ?? 30)) return null
  return excelDateSerial(year, month, day)
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
}

function escapeForRegex(char: string): string {
  return char.replace(/[.*+?^${}()|[\]\\\-/]/g, '\\$&')
}

/// Parses a same-sheet A1 destination ("B2"); null when the text is not a
/// plain cell reference (cross-sheet destinations are not offered).
export function parseDestinationCell(reference: string): { row: number; column: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(reference.trim())
  if (!match) return null
  const column = lettersToColumnIndex(match[1] ?? '')
  const row = Number(match[2])
  if (row < 1 || row > 1_048_576) return null
  return { row: row - 1, column }
}

function lettersToColumnIndex(letters: string): number {
  let column = 0
  for (const char of letters.toUpperCase()) column = column * 26 + (char.charCodeAt(0) - 64)
  return column - 1
}

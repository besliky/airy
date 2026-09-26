/**
 * Excel table-name rules and structured-reference text rewrites (PAR-202).
 *
 * Pure string algebra shared by the renderer (live validation/messages) and
 * the gateway (fail-closed rewrites inside the saved package). The gateway
 * never trusts the renderer: every rule here is re-checked there.
 */

export const SHEET_MAX_ROWS = 1_048_576
export const SHEET_MAX_COLUMNS = 16_384

/// Geometry a rewrite needs: 0-based inclusive bounds plus the band sizes the
/// structured-reference selectors resolve against.
export interface TableRefGeometry {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  readonly headerRowCount: number
  readonly totalsRowCount: number
}

/// The token a structured reference uses for a table: letters, digits,
/// underscore, period, backslash; must not start with a digit; no spaces.
const TABLE_NAME_PATTERN = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u

/// Characters that may not touch a bare table token inside a formula.
const BOUNDARY = '[^\\p{L}\\p{N}_.\\\\\'"]'
/// End boundary for the bare whole-table form: also excludes "(", which
/// turns the token into a function call instead of a table reference.
const BOUNDARY_END = '[^\\p{L}\\p{N}_.\\\\\'"(]'

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/// A1 column label → 0-based index ('A' → 0). Returns -1 for bad labels.
function labelToIndex(label: string): number {
  let index = 0
  for (const char of label) {
    const upper = char.toUpperCase()
    if (upper < 'A' || upper > 'Z') return -1
    index = index * 26 + (upper.charCodeAt(0) - 64)
  }
  return index - 1
}

/// English rule text for a rejected table name, or null when Excel would
/// accept it. The renderer shows it verbatim; the gateway throws it.
export function tableNameError(name: string): string | null {
  if (name.length === 0) return 'Table names cannot be empty.'
  if (name.length > 255) return 'Table names are limited to 255 characters.'
  // Single-letter R/C are reserved in Excel's name space.
  if (/^[CcRr]$/.test(name)) return `"${name}" is reserved by Excel and cannot be a table name.`
  if (!TABLE_NAME_PATTERN.test(name)) {
    return 'Table names must start with a letter, underscore, or backslash and may contain only letters, digits, periods, underscores, and backslashes — no spaces.'
  }
  // Cell-reference lookalikes are rejected: A1-style within the sheet's
  // column range (Excel blocks "TAB1" but accepts column-less "ZZZ99"-style
  // names beyond XFD).
  const a1 = /^([A-Za-z]{1,3})([0-9]+)$/.exec(name)
  const columnIndex = a1?.[1] ? labelToIndex(a1[1]) : -1
  if (a1 && columnIndex >= 0 && columnIndex < SHEET_MAX_COLUMNS) {
    return `"${name}" looks like a cell reference, which Excel does not allow as a table name.`
  }
  // R1C1-style lookalikes: R, R1, R1C1, C1 …
  if (/^[Rr]([0-9]+)?([Cc]([0-9]+)?)?$/.test(name) || /^[Cc]([0-9]+)?$/.test(name)) {
    return `"${name}" looks like an R1C1 reference, which Excel does not allow as a table name.`
  }
  return null
}

/// True when a structured reference must quote the table name (names with
/// characters outside the identifier set, e.g. written by other producers).
export function tableNameNeedsQuotes(name: string): boolean {
  return !TABLE_NAME_PATTERN.test(name)
}

function xmlEscapeText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function xmlEscapeQuoted(text: string): string {
  return `'${xmlEscapeText(text).replace(/'/g, "''")}'`
}

/// Skips a double-quoted string literal ("" is an escaped quote) or a
/// single-quoted sheet name ('' is an escaped apostrophe) starting at
/// `index`; returns the index just past the closing quote.
function skipQuoted(text: string, index: number, quote: '"' | "'"): number {
  index += 1
  while (index < text.length) {
    if (text[index] === quote) {
      if (text[index + 1] === quote) index += 2
      else return index + 1
    } else index += 1
  }
  return index
}

interface ScanState {
  result: string
  segmentStart: number
  index: number
}

/// Shared scanner: walks the formula outside string literals and quoted
/// sheet names, letting `tryMatch` rewrite tokens at the current position.
function scanFormula(
  formula: string,
  tryMatch: (formula: string, index: number, state: ScanState) => boolean,
): string {
  const state: ScanState = { result: '', segmentStart: 0, index: 0 }
  while (state.index < formula.length) {
    const char = formula[state.index]
    if (char === '"') {
      state.index = skipQuoted(formula, state.index, '"')
      continue
    }
    if (char === "'") {
      const end = skipQuoted(formula, state.index, "'")
      // A quote followed by "!" qualifies a sheet name — the whole region is
      // off limits. Any other quote wraps a quoted table token (or stray
      // text), which the matchers below may rewrite.
      if (formula[end] === '!') {
        state.index = end
        continue
      }
      if (tryMatch(formula, state.index, state)) continue
      state.index += 1
      continue
    }
    if (tryMatch(formula, state.index, state)) continue
    state.index += 1
  }
  return state.result + formula.slice(state.segmentStart)
}

/// Rewrites every structured reference to table `from` in a formula body
/// (the XML-escaped text between <f> tags) into a reference to `to`. Handles
/// the bare token (`Users[Col]`, whole-table `Users`) and the quoted token
/// (`'Old Name'[Col]`, apostrophes doubled). String literals ("…") and
/// quoted sheet names ('…'!) are skipped so cell text is never touched.
export function renameTableInFormulaText(formula: string, from: string, to: string): string {
  if (from.length === 0) return formula
  const bareFrom = xmlEscapeText(from)
  const quotedFrom = xmlEscapeQuoted(from)
  const bareTo = tableNameNeedsQuotes(to) ? xmlEscapeQuoted(to) : xmlEscapeText(to)
  const quotedTo = xmlEscapeQuoted(to)
  // The quoted form must be followed by a selector bracket; the bare form
  // covers both `Name[Col]` and the whole-table reference. A token followed
  // directly by "(" is a function call, never a table.
  const patterns = [
    {
      source: new RegExp(`(?:${BOUNDARY}|^)${escapeRegExp(quotedFrom)}(?=\\[)`, 'giu'),
      selector: true,
    },
    {
      source: new RegExp(`(?:${BOUNDARY}|^)${escapeRegExp(bareFrom)}(?=\\[)`, 'giu'),
      selector: true,
    },
    {
      source: new RegExp(`(?:${BOUNDARY}|^)${escapeRegExp(bareFrom)}(?=${BOUNDARY_END}|$)`, 'giu'),
      selector: false,
    },
  ]
  return scanFormula(formula, (text, index, state) => {
    const rest = text.slice(index)
    for (const pattern of patterns) {
      pattern.source.lastIndex = 0
      const match = pattern.source.exec(rest)
      if (!match || match.index !== 0) continue
      // The boundary alternative keeps exactly one leading character; a
      // leading apostrophe can only be the quoted token (the boundary class
      // excludes apostrophes), which decides bare vs quoted output.
      const replacement = match[0].startsWith("'") ? quotedTo : bareTo
      // A zero-width prefix ("^" alternative) is only honest at the very
      // start of the formula — otherwise the slice began mid-identifier
      // ("MyUsers" would rewrite its "Users" tail).
      const tokenOnly = match[0] === (match[0].startsWith("'") ? quotedFrom : bareFrom)
      if (tokenOnly && index !== 0) continue
      const swapped = match[0].replace(
        new RegExp(`${escapeRegExp(quotedFrom)}|${escapeRegExp(bareFrom)}`, 'giu'),
        () => replacement,
      )
      state.result += text.slice(state.segmentStart, index) + swapped
      // Selector forms ("Name[…]") skip their whole bracket group (kept
      // verbatim), so column names inside it never read as table tokens.
      const groupEnd = pattern.selector ? findSelectorEnd(text, index + match[0].length) : null
      if (groupEnd !== null) {
        state.result += text.slice(index + match[0].length, groupEnd)
        state.index = groupEnd
      } else {
        state.index = index + match[0].length
      }
      state.segmentStart = state.index
      return true
    }
    return false
  })
}

/// Index just past the structured-reference bracket group opening at
/// `openIndex` (which must point at '['), or null when the group never
/// closes. Plain depth counting: column names containing "]" are beyond
/// this heuristic's reach (they stay untouched instead of being rewritten).
function findSelectorEnd(text: string, openIndex: number): number | null {
  let depth = 0
  let index = openIndex
  while (index < text.length) {
    if (text[index] === '[') depth += 1
    else if (text[index] === ']') {
      depth -= 1
      index += 1
      if (depth === 0) return index
      continue
    }
    index += 1
  }
  return null
}

type RefBand = 'all' | 'data' | 'headers' | 'totals' | 'thisRow'

interface RefSelection {
  /// Row bands the selector unions: a single band for plain selectors, or
  /// several special items ("[[#Data],[#Totals],[Col]]").
  readonly bands: readonly RefBand[]
  /// Column span resolved from names; null = all table columns.
  readonly startColumn: number | null
  readonly endColumn: number | null
}

/// Decodes structured-reference escapes inside a bracket token.
function decodeRefToken(token: string): string {
  return token.replace(/''/g, "'").replace(/'#/g, '#').replace(/\[\[/g, '[').replace(/\]\]/g, ']')
}

/// Parses the bracket selector following a table name starting at `start`
/// (which must point at '['). Returns the selection and the index just past
/// the closing bracket, or null for a malformed selector.
function parseRefSelector(
  text: string,
  start: number,
  columns: readonly string[],
): { selection: RefSelection; end: number } | null {
  let depth = 0
  let index = start
  let token = ''
  const tokens: string[] = []
  while (index < text.length) {
    const char = text[index]
    if (char === '[') {
      depth += 1
      if (depth > 3) return null
      index += 1
      continue
    }
    if (char === ']') {
      depth -= 1
      index += 1
      if (depth === 0) {
        // The group's final close carries the last token — but nested
        // groups ("[[A]:[B]]") close once more with nothing left.
        if (token !== '') tokens.push(token)
        break
      }
      tokens.push(token)
      token = ''
      // A ":" or "," separator continues the group; the following "[" is
      // consumed by the depth branch above on the next pass.
      const separator = /^\s*[:,]/.exec(text.slice(index))
      if (separator) index += separator[0].length
      continue
    }
    token += char
    index += 1
  }
  if (depth !== 0 || tokens.length === 0) return null

  const decoded = tokens.map(decodeRefToken)
  // Every special item in the group unions its band ("[[#Data],[#Totals],[Col]]"
  // spans data plus totals); Excel's Convert to Range folds contiguous bands
  // into one A1 range.
  const bands = new Set<RefBand>()
  for (const entry of decoded) {
    const value = entry.trim().toLowerCase()
    if (value === '#all') bands.add('all')
    else if (value === '#headers') bands.add('headers')
    else if (value === '#data') bands.add('data')
    else if (value === '#totals') bands.add('totals')
    else if (value === '#this row') bands.add('thisRow')
  }
  // " [@Col]" is the legacy spelling of "[[#This Row],[Col]]".
  if (bands.size === 0) {
    bands.add(decoded.some((entry) => entry.trim().startsWith('@')) ? 'thisRow' : 'data')
  }
  const columnTokens = decoded
    .filter((entry) => !/^#\w/i.test(entry.trim()))
    .map((entry) => entry.trim().replace(/^@/, ''))
    .filter((entry) => entry !== '')
  if (columnTokens.length === 0) {
    return { selection: { bands: [...bands], startColumn: null, endColumn: null }, end: index }
  }
  const indexes: number[] = []
  for (const columnToken of columnTokens) {
    for (const part of columnToken.split(':')) {
      const needle = part.trim().toLowerCase()
      const found = columns.findIndex((column) => column.toLowerCase() === needle)
      if (found < 0) return null
      indexes.push(found)
    }
  }
  return {
    selection: {
      bands: [...bands],
      startColumn: Math.min(...indexes),
      endColumn: Math.max(...indexes),
    },
    end: index,
  }
}

/// Resolves a selection to an absolute 0-based row range, or null when a
/// selected band is empty (totals without a totals row), "this row" falls
/// outside the data body, or the union has a gap (headers and totals without
/// data) — such unions have no single A1 equivalent.
function selectionRows(
  selection: RefSelection,
  table: TableRefGeometry,
  formulaRow: number,
): { start: number; end: number } | null {
  const ranges: { start: number; end: number }[] = []
  for (const band of selection.bands) {
    const rows = singleBandRows(band, table, formulaRow)
    if (rows === null) return null
    ranges.push(rows)
  }
  ranges.sort((a, b) => a.start - b.start || a.end - b.end)
  let end = ranges[0]!.end
  for (const range of ranges.slice(1)) {
    if (range.start > end + 1) return null
    end = Math.max(end, range.end)
  }
  return { start: ranges[0]!.start, end }
}

/// Row range of one band, or null when the band resolves to nothing.
function singleBandRows(
  band: RefBand,
  table: TableRefGeometry,
  formulaRow: number,
): { start: number; end: number } | null {
  const headerEnd = table.startRow + table.headerRowCount - 1
  const totalsStart = table.endRow - table.totalsRowCount + 1
  switch (band) {
    case 'all':
      return { start: table.startRow, end: table.endRow }
    case 'headers':
      return table.headerRowCount > 0 ? { start: table.startRow, end: headerEnd } : null
    case 'totals':
      return table.totalsRowCount > 0 ? { start: totalsStart, end: table.endRow } : null
    case 'data':
      return { start: headerEnd + 1, end: totalsStart - 1 }
    case 'thisRow':
      return formulaRow >= headerEnd + 1 && formulaRow <= totalsStart - 1
        ? { start: formulaRow, end: formulaRow }
        : null
  }
}

function cellLabel(row: number, column: number): string {
  let label = ''
  let index = column + 1
  while (index > 0) {
    const remainder = (index - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    index = Math.floor((index - 1) / 26)
  }
  return `${label}${row + 1}`
}

function rangeText(
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
): string {
  const start = cellLabel(startRow, startColumn)
  if (startRow === endRow && startColumn === endColumn) return start
  return `${start}:${cellLabel(endRow, endColumn)}`
}

export interface TableA1Rewrite {
  readonly name: string
  readonly geometry: TableRefGeometry
  readonly columns: readonly string[]
  readonly sheetName: string
}

/// Excel's Convert to Range rule (Microsoft: "when you convert a table to a
/// range, all cell references change to their equivalent A1 style
/// references"): every structured reference to the table becomes its A1
/// equivalent. Same-sheet formulas stay unqualified, other sheets get a
/// sheet-qualified reference, and "this row" resolves against the formula
/// cell's own row. Unresolvable selectors become #REF!.
export function tableRefsToA1InFormulaText(
  formula: string,
  table: TableA1Rewrite,
  formulaRow: number,
  formulaSheetName: string,
): string {
  const bareName = xmlEscapeText(table.name)
  const quotedName = xmlEscapeQuoted(table.name)
  const qualification =
    formulaSheetName === table.sheetName ? '' : `${sheetQualify(table.sheetName)}!`
  // Selector form: boundary + name + bracket; whole-table form: boundary +
  // name + boundary/end. The consuming boundary character keeps the scan
  // honest inside longer identifiers ("MyUsers" never yields its "Users").
  const patterns = tableNameNeedsQuotes(table.name)
    ? [
        {
          source: new RegExp(`(?:${BOUNDARY}|^)${escapeRegExp(quotedName)}(?=\\[)`, 'giu'),
          selector: true,
        },
        {
          source: new RegExp(
            `(?:${BOUNDARY}|^)${escapeRegExp(quotedName)}(?=${BOUNDARY_END}|$)`,
            'giu',
          ),
          selector: false,
        },
      ]
    : [
        {
          source: new RegExp(`(?:${BOUNDARY}|^)${escapeRegExp(bareName)}(?=\\[)`, 'giu'),
          selector: true,
        },
        {
          source: new RegExp(
            `(?:${BOUNDARY}|^)${escapeRegExp(bareName)}(?=${BOUNDARY_END}|$)`,
            'giu',
          ),
          selector: false,
        },
      ]
  const nameLengths = tableNameNeedsQuotes(table.name)
    ? [quotedName.length, quotedName.length]
    : [bareName.length, bareName.length]
  return scanFormula(formula, (text, index, state) => {
    const rest = text.slice(index)
    for (let entry = 0; entry < patterns.length; entry += 1) {
      const pattern = patterns[entry]!
      pattern.source.lastIndex = 0
      const match = pattern.source.exec(rest)
      if (!match || match.index !== 0) continue
      // Zero-width prefix is only honest at the formula's start (see
      // renameTableInFormulaText).
      if (match[0].length === (nameLengths[entry] ?? 0) && index !== 0) continue
      const prefix = match[0].slice(0, match[0].length - (nameLengths[entry] ?? 0))
      if (entry === 0) {
        const openIndex = index + match[0].length
        const selector = parseRefSelector(text, openIndex, table.columns)
        const text2 =
          selector === null
            ? '#REF!'
            : selectionText(selector.selection, table.geometry, formulaRow, qualification)
        state.result += text.slice(state.segmentStart, index) + prefix + text2
        // A malformed or unknown selector still consumes its whole group so
        // the leftovers never rescan as tokens.
        state.index =
          selector === null ? (findSelectorEnd(text, openIndex) ?? openIndex) : selector.end
      } else {
        const rows = selectionRows(
          { bands: ['data'], startColumn: null, endColumn: null },
          table.geometry,
          formulaRow,
        )
        const body =
          rows === null
            ? '#REF!'
            : rangeText(rows.start, rows.end, table.geometry.startColumn, table.geometry.endColumn)
        state.result += text.slice(state.segmentStart, index) + prefix + qualification + body
        state.index = index + match[0].length
      }
      state.segmentStart = state.index
      return true
    }
    return false
  })
}

function selectionText(
  selection: RefSelection,
  table: TableRefGeometry,
  formulaRow: number,
  qualification: string,
): string {
  const rows = selectionRows(selection, table, formulaRow)
  if (rows === null) return '#REF!'
  const startColumn =
    selection.startColumn === null ? table.startColumn : table.startColumn + selection.startColumn
  const endColumn =
    selection.endColumn === null ? table.endColumn : table.startColumn + selection.endColumn
  return qualification + rangeText(rows.start, rows.end, startColumn, endColumn)
}

/// Excel quotes sheet names in references only when they contain characters
/// outside the plain identifier set.
function sheetQualify(name: string): string {
  return /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

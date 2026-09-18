import type { Command, EditorState } from '@tiptap/pm/state'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import { isInTable, selectedRect } from '@tiptap/pm/tables'

/**
 * Word's Home ▸ Sort: reorder whole table rows by a column key, or reorder the
 * selected paragraphs as lines. Rows and paragraphs travel as whole nodes
 * (formatting, row heights, revisions and save anchors ride along), so a sort
 * is one replace transaction — a single undo step.
 */

export type SortFieldType = 'text' | 'number' | 'date'

export interface SortLevel {
  /** zero-based displayed column (table scope); ignored by the paragraph sort */
  column: number
  type: SortFieldType
  descending: boolean
}

export interface SortOptions {
  /** first entry is the primary key, the rest are Word's "Then by" levels */
  levels: SortLevel[]
  /** tables: keep row 0 (the header) out of the sort */
  headerRow: boolean
}

/** What the Sort dialog would act on for the current selection. */
export interface SortScopeTable {
  kind: 'table'
  /** grid width in displayed columns (colspans expanded) */
  columnCount: number
  rowCount: number
  /** non-empty texts of row 0 cells, for Word-style column names */
  headerLabels: string[]
  /** default header-row state: a repeating header or a docTableHeader row */
  headerRow: boolean
  /** vertical merges (rowspan) anywhere: refused, see tableHasVerticalMerge */
  hasVerticalMerge: boolean
}

export interface SortScopeParagraphs {
  kind: 'paragraphs'
  /** top-level blocks the selection touches (all paragraph-like; else null scope) */
  count: number
}

export type SortScope = SortScopeTable | SortScopeParagraphs | null

/** the table the selection is in (caret or whole-table NodeSelection, like the context menu) */
function tableTarget(state: EditorState): { table: PmNode; pos: number } | null {
  const selection = state.selection
  if (selection instanceof NodeSelection && selection.node.type.name === 'docTable') {
    return { table: selection.node, pos: selection.from }
  }
  if (!isInTable(state)) return null
  try {
    const rect = selectedRect(state)
    const pos = rect.tableStart - 1
    const node = state.doc.nodeAt(pos)
    return node?.type.name === 'docTable' ? { table: node, pos } : null
  } catch {
    return null
  }
}

/** grid width (displayed columns) of a rowspan-1 table row */
function rowColumnCount(row: PmNode): number {
  let columns = 0
  row.forEach((cell) => {
    columns += Math.max(1, Number(cell.attrs.colspan) || 1)
  })
  return columns
}

/**
 * True when any cell spans rows. Word keeps vertical merges intact by refusing
 * to sort (moving a row out of a merged run would orphan the continuation
 * cells); the rowspan model has the same problem, so sorting is declined.
 */
export function tableHasVerticalMerge(table: PmNode): boolean {
  let merged = false
  table.forEach((row) => {
    row.forEach((cell) => {
      if (Math.max(1, Number(cell.attrs.rowspan) || 1) > 1) merged = true
    })
  })
  return merged
}

/** displayed-column index -> XML cell covering that grid slot (rowspan-1 rows only) */
function keyCell(
  row: PmNode,
  displayedColumn: number,
  columnCount: number,
  rtl: boolean,
): PmNode | null {
  // w:bidiVisual renders the first XML cell on the right: displayed column 0 is
  // the last XML grid slot
  const slot = rtl ? columnCount - 1 - displayedColumn : displayedColumn
  let cursor = 0
  let hit: PmNode | null = null
  row.forEach((cell) => {
    const span = Math.max(1, Number(cell.attrs.colspan) || 1)
    if (hit === null && slot >= cursor && slot < cursor + span) hit = cell
    cursor += span
  })
  return hit
}

function cellText(cell: PmNode | null): string {
  if (!cell) return ''
  return cell.textContent.replace(/\u00a0/g, ' ').trim()
}

/* ================= key parsing ================= */

/** Word-style number reading: currency signs, percent, spaces, thousands
 *  separators and parenthesized negatives are tolerated; null = not a number
 *  (sorts after numeric entries). */
export function parseSortNumber(text: string): number | null {
  let s = text.replace(/[\s\u00a0\u202f\u2009]/g, '')
  if (!s) return null
  const negative = /^\((.*)\)$/.exec(s)
  if (negative) s = negative[1]
  s = s.replace(/[$€£¥₹₽%°]/g, '')
  if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s))
    s = s.replace(/,/g, '') // thousands
  else if (/^\d+,\d+$/.test(s))
    s = s.replace(',', '.') // locale decimal comma
  else if (/^(\d{1,3}(\.\d{3})+),(\d+)$/.test(s)) {
    // 1.234,56: dot thousands + comma decimal
    s = s.replace(/\./g, '').replace(',', '.')
  }
  const value = Number(s)
  if (!Number.isFinite(value)) return null
  return negative ? -value : value
}

const MONTHS: Record<string, number> = {
  jan: 1,
  january: 1,
  feb: 2,
  february: 2,
  mar: 3,
  march: 3,
  apr: 4,
  april: 4,
  may: 5,
  jun: 6,
  june: 6,
  jul: 7,
  july: 7,
  aug: 8,
  august: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  october: 10,
  nov: 11,
  november: 11,
  dec: 12,
  december: 12,
}

const dateAt = (y: number, m: number, d: number, hh = 0, mm = 0): number | null => {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null
  if (hh > 23 || mm > 59) return null
  // reject rollover readings Word treats as text: Feb 30, Apr 31, 25:00 —
  // Date.UTC would silently normalize them into the next month/day
  const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0
  const lengthOfMonth = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]!
  if (d > lengthOfMonth) return null
  // Date.UTC maps years 0-99 to 1900 + y (its two-digit rule) — re-pin the
  // written year so "0066-05-05" reads as year 66, not 1966
  if (y >= 0 && y <= 99) {
    const utc = new Date(Date.UTC(y, m - 1, d, hh, mm))
    utc.setUTCFullYear(y)
    return utc.getTime()
  }
  return Date.UTC(y, m - 1, d, hh, mm)
}

/** Word-style date reading for common formats (ISO, numeric, "12 Jan 2026",
 *  "Jan 12, 2026", CJK Y/M/D order). Ambiguous d/m vs m/d: the side that cannot be
 *  a month forces the reading, otherwise day-first; null = not a date. */
export function parseSortDate(text: string): number | null {
  const s = text.trim()
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s](\d{1,2}):(\d{2}))?/.exec(s)
  if (m) return dateAt(+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0))
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?:\s(\d{1,2}):(\d{2}))?$/.exec(s)
  if (m) {
    const year = +m[3] + (m[3].length === 2 ? (year2k(+m[3]) ? 2000 : 1900) : 0)
    const a = +m[1]
    const b = +m[2]
    // whichever side cannot be a month forces the reading; both fitting
    // defaults day-first (the more common global convention)
    let day = a
    let month = b
    if (a > 12 && b > 12) return null
    if (a > 12) month = b
    else if (b > 12) {
      month = a
      day = b
    }
    return dateAt(year, month, day, +(m[4] ?? 0), +(m[5] ?? 0))
  }
  m = /^(\d{1,2})\s*([A-Za-z]+)\.?\s*(\d{4})$/.exec(s)
  if (m) return dateAt(+m[3], MONTHS[m[2].toLowerCase()] ?? 0, +m[1])
  m = /^([A-Za-z]+)\.?\s*(\d{1,2}),?\s*(\d{4})$/.exec(s)
  if (m) return dateAt(+m[3], MONTHS[m[1].toLowerCase()] ?? 0, +m[2])
  m = /^(\d{4})年(\d{1,2})月(\d{1,2})日?$/.exec(s)
  if (m) return dateAt(+m[1], +m[2], +m[3])
  return null
}

/** two-digit years: 00-39 -> 2000s, 40-99 -> 1900s (the Excel pivot) */
function year2k(two: number): boolean {
  return two <= 39
}

function compareByType(a: string, b: string, type: SortFieldType): number {
  if (type === 'number') {
    const x = parseSortNumber(a)
    const y = parseSortNumber(b)
    // unparseable entries sort after numeric ones ascending; `descending`
    // negates the comparison wholesale (compareLevels), so they land BEFORE
    // numbers then — Word/Excel comparators order text above numbers in a
    // descending numeric column too (BUG-744: this is the intended behavior)
    if (x === null && y === null) return 0
    if (x === null) return 1
    if (y === null) return -1
    return x - y
  }
  if (type === 'date') {
    const x = parseSortDate(a)
    const y = parseSortDate(b)
    if (x === null && y === null) return 0
    if (x === null) return 1
    if (y === null) return -1
    return x - y
  }
  return a.localeCompare(b, undefined, { numeric: true })
}

/** primary + then-by comparison of per-level key lists (positional, not by column) */
function compareLevels(
  a: readonly string[],
  b: readonly string[],
  levels: readonly SortLevel[],
): number {
  for (let i = 0; i < levels.length; i++) {
    const order = compareByType(a[i] ?? '', b[i] ?? '', levels[i].type)
    if (order !== 0) return levels[i].descending ? -order : order
  }
  return 0
}

/* ================= scope ================= */

/** What the sort dialog should show for the current selection. */
export function sortScope(state: EditorState): SortScope {
  const target = tableTarget(state)
  if (target) {
    const { table } = target
    const rows: PmNode[] = []
    table.forEach((row) => rows.push(row))
    const columnCount = rows.length ? rowColumnCount(rows[0]) : 0
    const headerRow =
      rows.length > 0 &&
      (rows[0].attrs.repeatHeader === true ||
        (rows[0].firstChild?.type.name ?? '') === 'docTableHeader')
    return {
      kind: 'table',
      columnCount,
      rowCount: rows.length,
      headerLabels: rows.length
        ? rowColumnTexts(rows[0], columnCount, table.attrs.bidiVisual === true)
        : [],
      headerRow,
      hasVerticalMerge: tableHasVerticalMerge(table),
    }
  }
  const range = selectedParagraphRange(state.doc, state.selection.from, state.selection.to)
  if (!range) return null
  return { kind: 'paragraphs', count: range.blocks.length }
}

/** texts of row 0 cells by displayed column, for Word-style column names */
function rowColumnTexts(row: PmNode, columnCount: number, rtl: boolean): string[] {
  const labels: string[] = []
  for (let column = 0; column < columnCount; column++) {
    labels.push(cellText(keyCell(row, column, columnCount, rtl)))
  }
  return labels
}

/** top-level paragraph-like blocks the selection touches (null when the range
 *  holds anything else — tables, content-control shells — or spans < 2 blocks) */
interface ParagraphRange {
  blocks: PmNode[]
  start: number
  end: number
}

const PARAGRAPH_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

function selectedParagraphRange(doc: PmNode, from: number, to: number): ParagraphRange | null {
  if (from === to) return null
  // top-level block indexes (index(0)); deeper nesting (table cells) fails the
  // paragraph-like check below, and tables sort through the table path instead
  const $from = doc.resolve(from)
  const $to = doc.resolve(to)
  const first = $from.index(0)
  // a selection ending exactly at a block start does not include that block
  const last = $to.parentOffset === 0 && $to.index(0) > first ? $to.index(0) - 1 : $to.index(0)
  if (last - first < 1) return null
  const blocks: PmNode[] = []
  let start = 0
  for (let i = 0; i < first; i++) start += doc.child(i).nodeSize
  let end = start
  for (let i = first; i <= last; i++) {
    const block = doc.child(i)
    // sdt group members carry shell bytes that must stay in order
    if (!PARAGRAPH_TYPES.has(block.type.name) || block.attrs.sdtShell) return null
    blocks.push(block)
    end += block.nodeSize
  }
  return { blocks, start, end }
}

/* ================= commands ================= */

/** Sort the selected table's rows by the given levels; declines vertical merges. */
export function sortTableRows(options: SortOptions): Command {
  return (state, dispatch) => {
    const target = tableTarget(state)
    if (!target) return false
    const { table, pos } = target
    if (tableHasVerticalMerge(table)) return false
    const rows: PmNode[] = []
    table.forEach((row) => rows.push(row))
    const headerCount = options.headerRow ? 1 : 0
    if (rows.length - headerCount < 2) return false
    const levels = options.levels
    if (!levels.length) return false
    const columnCount = rowColumnCount(rows[0])
    const keyed = rows.slice(headerCount).map((row) => ({
      row,
      keys: levels.map((l) =>
        cellText(keyCell(row, l.column, columnCount, table.attrs.bidiVisual === true)),
      ),
    }))
    keyed.sort((a, b) => compareLevels(a.keys, b.keys, levels))
    const sorted = [...rows.slice(0, headerCount), ...keyed.map((k) => k.row)]
    if (sorted.every((row, i) => row === rows[i])) return false
    if (dispatch) {
      const tr = state.tr
      tr.replaceWith(pos + 1, pos + table.nodeSize - 1, sorted)
      tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 2)))
      dispatch(tr)
    }
    return true
  }
}

/** Sort the selected top-level paragraphs by their text. */
export function sortSelectedParagraphs(levels: SortLevel[]): Command {
  const level: SortLevel = levels[0] ?? {
    column: 0,
    type: 'text',
    descending: false,
  }
  return (state, dispatch) => {
    const range = selectedParagraphRange(state.doc, state.selection.from, state.selection.to)
    if (!range) return false
    const keyed = range.blocks.map((block) => block.textContent.replace(/\u00a0/g, ' ').trim())
    const order = keyed
      .map((text, index) => ({ text, index }))
      .sort((a, b) => {
        const result = compareByType(a.text, b.text, level.type)
        return (level.descending ? -result : result) || a.index - b.index
      })
      .map((entry) => entry.index)
    if (order.every((i, position) => i === position)) return false
    if (dispatch) {
      const tr = state.tr
      tr.replaceWith(
        range.start,
        range.end,
        order.map((i) => range.blocks[i]),
      )
      tr.setSelection(TextSelection.near(tr.doc.resolve(range.end)))
      dispatch(tr)
    }
    return true
  }
}

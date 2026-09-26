import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Mark as PmMark } from '@tiptap/pm/model'

import {
  evaluateFormulaInGrid,
  proposeTableFormula as proposeFromGrid,
  type TableCell,
  type TableModel,
} from '@airy-office/docx-engine'

import { TRACK_IGNORE, type FieldCacheJob } from './revisions'

/**
 * Table formula fields (=SUM(ABOVE)…), Word's Table Layout → Formula.
 * A formula lives in one cell paragraph as a text run carrying the
 * `tableFormula` mark; the run text is the cached result (Word shows the same
 * until F9). This module builds the physical grid of a native table, inserts
 * formulas and collects the F9 update jobs.
 */

const CELL_TYPES = new Set(['docTableCell', 'docTableHeader'])

interface PhysicalCell {
  node: PmNode
  /** absolute position of the cell node */
  pos: number
  row: number
  col: number
}

export interface TableGridInfo {
  table: PmNode
  tablePos: number
  /** physical grid of cell texts (rowspan/colspan expanded), [row][col] */
  texts: string[][]
  /** physical row indexes whose rows repeat as headers (w:tblHeader) */
  headerRows: Set<number>
  /** every real (non-continuation) cell with its physical row/col */
  cells: PhysicalCell[]
}

/**
 * Visible text of a cell for grid math. Nested docTable subtrees do not
 * contribute: Word treats a table-holding cell as non-numeric, so an outer
 * formula must never read the nested cells' (concatenated) text as a number.
 */
function cellOwnText(cell: PmNode): string {
  let text = ''
  cell.descendants((n) => {
    if (n.type.name === 'docTable') return false
    if (n.isText) text += n.text
    return true
  })
  return text
}

/** Build the physical grid of a native docTable (mirrors pmTableToModel's spans). */
export function pmTableGrid(table: PmNode, tablePos: number): TableGridInfo {
  const headerRows = new Set<number>()
  const cells: PhysicalCell[] = []
  const texts: string[][] = []
  // rowspan continuations carried into following rows: col → { remaining, span }
  const active = new Map<number, { remaining: number; span: number }>()
  let rowIndex = -1
  table.forEach((rowNode, rowOffset) => {
    rowIndex++
    const rowPos = tablePos + 1 + rowOffset
    if (rowNode.attrs?.repeatHeader) headerRows.add(rowIndex)
    texts[rowIndex] = []
    const occupied = new Set<number>()
    for (const [col, span] of active) {
      for (let i = 0; i < span.span; i++) {
        texts[rowIndex][col + i] = ''
        occupied.add(col + i)
      }
    }
    // direct cell children only: a nested table's cells belong to the nested
    // table's own grid, never to this one (BUG-1757)
    const rowCells: Array<{ node: PmNode; pos: number }> = []
    rowNode.forEach((cellNode, cellOffset) => {
      if (CELL_TYPES.has(cellNode.type.name)) {
        rowCells.push({ node: cellNode, pos: rowPos + 1 + cellOffset })
      }
    })
    let cursor = 0
    const added = new Map<number, { remaining: number; span: number }>()
    for (const cell of rowCells) {
      while (occupied.has(cursor)) cursor++
      const colspan = Math.max(1, Number(cell.node.attrs?.colspan) || 1)
      const rowspan = Math.max(1, Number(cell.node.attrs?.rowspan) || 1)
      const text = cellOwnText(cell.node)
      for (let i = 0; i < colspan; i++) texts[rowIndex][cursor + i] = i === 0 ? text : ''
      cells.push({ node: cell.node, pos: cell.pos, row: rowIndex, col: cursor })
      if (rowspan > 1) added.set(cursor, { remaining: rowspan - 1, span: colspan })
      for (let i = 0; i < colspan; i++) occupied.add(cursor + i)
      cursor += colspan
    }
    const next = new Map<number, { remaining: number; span: number }>()
    for (const [col, span] of active) {
      if (span.remaining > 1) next.set(col, { ...span, remaining: span.remaining - 1 })
    }
    for (const [col, span] of added) next.set(col, span)
    active.clear()
    for (const [k, v] of next) active.set(k, v)
  })
  return { table, tablePos, texts, headerRows, cells }
}

const gridOf = (info: TableGridInfo) => ({
  texts: info.texts,
  isHeaderRow: (r: number) => info.headerRows.has(r),
})

export interface FormulaCell {
  /** the cell's own text cache entry (grid row/col) */
  row: number
  col: number
  grid: TableGridInfo
}

/** Locate the table cell containing the caret, with its physical grid placement. */
export function selectedFormulaCell(editor: Editor): FormulaCell | null {
  const { state } = editor
  const selection = state.selection
  for (const $pos of [selection.$from, selection.$anchor]) {
    for (let d = $pos.depth; d > 0; d--) {
      const node = $pos.node(d)
      if (!CELL_TYPES.has(node.type.name)) continue
      const tableDepth = d - 2
      if (tableDepth < 0) break
      const tableNode = $pos.node(tableDepth)
      if (tableNode.type.name !== 'docTable') break
      const tablePos = $pos.before(tableDepth)
      const grid = pmTableGrid(tableNode, tablePos)
      const cellPos = $pos.before(d)
      const cell = grid.cells.find((c) => c.pos === cellPos)
      if (!cell) return null
      return { row: cell.row, col: cell.col, grid }
    }
  }
  return null
}

/** Word's dialog prefill: =SUM(ABOVE) when the column above holds numbers, else =SUM(LEFT) */
export function proposeCellFormula(editor: Editor): string {
  const sel = selectedFormulaCell(editor)
  if (!sel) return '=SUM(ABOVE)'
  return proposeFromGrid(sel.grid.texts, sel.row, sel.col, (r) => sel.grid.headerRows.has(r))
}

/**
 * Insert a formula into the caret's cell paragraph (Word replaces the cell
 * paragraph content) and compute the cached result immediately.
 */
export function insertCellFormula(editor: Editor, instr: string): boolean {
  const sel = selectedFormulaCell(editor)
  if (!sel) return false
  const trimmed = instr.trim()
  if (!trimmed.startsWith('=')) return false
  const result = evaluateFormulaInGrid(trimmed, gridOf(sel.grid), sel.row, sel.col)
  const { state } = editor
  const $pos = state.selection.$from
  // the paragraph directly under the cell owns the formula (Word's placement)
  let paraDepth = -1
  for (let d = $pos.depth; d > 0; d--) {
    const name = $pos.node(d).type.name
    if (CELL_TYPES.has(name)) break
    if (name === 'docParagraph' || name === 'docListItem') paraDepth = d
  }
  if (paraDepth < 0) return false
  const mark = state.schema.marks.tableFormula?.create({ instr: trimmed })
  if (!mark) return false
  const from = $pos.before(paraDepth) + 1
  const to = $pos.after(paraDepth) - 1
  const textNode = state.schema.text(result, [mark])
  editor.view.dispatch(state.tr.replaceWith(from, Math.max(from, to), textNode).scrollIntoView())
  return true
}

/**
 * F9: recompute every table formula's cached result. Returns the cache jobs
 * (applied by applyFieldCaches in one TRACK_IGNORE transaction, like the
 * inline PAGE/REF refresh).
 */
export function collectTableFormulaJobs(editor: Editor): FieldCacheJob[] {
  const jobs: FieldCacheJob[] = []
  const visitCellText = (cell: PhysicalCell, grid: TableGridInfo) => {
    cell.node.descendants((n, off) => {
      if (n.type.name === 'docTable') return false // a nested table is its own grid
      if (!n.isText) return
      const mark = n.marks.find((m: PmMark) => m.type.name === 'tableFormula')
      if (!mark) return
      const instr = String(mark.attrs.instr)
      // a degenerate field without an instruction keeps its cached text: F9
      // must not write Word's "!Syntax Error" over it (BUG-1758)
      if (!instr.trim()) return
      const abs = cell.pos + 1 + off
      const next = evaluateFormulaInGrid(instr, gridOf(grid), cell.row, cell.col)
      if (next !== n.text)
        jobs.push({ from: abs, to: abs + n.nodeSize, text: next, marks: n.marks })
    })
  }
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docTable') return
    const grid = pmTableGrid(node, pos)
    for (const cell of grid.cells) visitCellText(cell, grid)
    // keep descending: a nested docTable is visited as its own grid so its
    // formulas resolve ABOVE/LEFT against its own rows/columns, like Word
  })
  return jobs
}

// ---- nested tables (docNestedTable atoms) ----
// A nested table is not PM content: it lives in the docNestedTable atom's
// `model` attribute (TableModel). Its formula caches therefore cannot ride the
// PM text jobs above — they are recomputed against the nested model's own grid
// and written back as model updates (BUG-1757).

interface ModelGrid {
  texts: string[][]
  isHeaderRow: (row: number) => boolean
}

/** visible text of a nested-model cell for grid math */
function nestedCellText(cell: TableCell): string {
  const paras = cell.richParas?.length
    ? cell.richParas.map((p) => p.runs.map((run) => run.text).join(''))
    : cell.paras
  return paras.join('\n')
}

/** physical texts grid of a nested TableModel: vMerge continuations and gridGap
 * placeholders contribute empty (non-numeric) text; colSpan widens its cell */
function nestedModelGrid(model: TableModel): ModelGrid {
  const texts: string[][] = []
  const repeats = model.repeatHeaderRows ?? []
  model.rows.forEach((row, r) => {
    const textsRow: string[] = []
    for (const cell of row) {
      textsRow.push(cell.gridGap || cell.vMerge === 'continue' ? '' : nestedCellText(cell))
      for (let i = 1; i < Math.max(1, cell.colSpan ?? 1); i++) textsRow.push('')
    }
    texts[r] = textsRow
  })
  return { texts, isHeaderRow: (row: number) => repeats[row] === true }
}

interface Refreshed<T> {
  value: T
  /** formula caches recomputed */
  count: number
}

/** recompute one nested cell's formula caches against the model grid; null = unchanged */
function refreshNestedCellFormulas(
  cell: TableCell,
  row: number,
  col: number,
  grid: ModelGrid,
): Refreshed<TableCell> | null {
  let count = 0
  let next = cell
  if (cell.richParas?.length) {
    let changed = false
    const richParas = cell.richParas.map((para) => {
      let paraChanged = false
      const runs = para.runs.map((run) => {
        if (run.formulaField === undefined) return run
        // a degenerate field without an instruction keeps its cached text (BUG-1758)
        if (!String(run.formulaField).trim()) return run
        const value = evaluateFormulaInGrid(String(run.formulaField), grid, row, col)
        if (value === run.text) return run
        paraChanged = true
        count++
        return { ...run, text: value }
      })
      if (!paraChanged) return para
      changed = true
      return { ...para, runs }
    })
    if (changed) {
      // keep the plain-text cache in step: the save diff compares cell paras
      const paras = cell.paras.map((text, p) =>
        richParas[p] === cell.richParas![p] ? text : richParas[p].runs.map((r) => r.text).join(''),
      )
      next = { ...cell, paras, richParas }
    }
  }
  if (cell.nestedTables?.length) {
    let nestedChanged = false
    const nestedTables = cell.nestedTables.map((nt) => {
      const rec = refreshNestedModelFormulas(nt)
      if (!rec) return nt
      count += rec.count
      nestedChanged = true
      return rec.value
    })
    if (nestedChanged) next = { ...next, nestedTables }
  }
  return next === cell ? null : { value: next, count }
}

/** recompute every formula of one nested TableModel on its own grid; null = unchanged */
function refreshNestedModelFormulas(model: TableModel): Refreshed<TableModel> | null {
  const grid = nestedModelGrid(model)
  let count = 0
  let changed = false
  const rows = model.rows.map((row, r) => {
    let cursor = 0
    return row.map((cell) => {
      const col = cursor
      cursor += Math.max(1, cell.colSpan ?? 1)
      if (cell.gridGap || cell.vMerge === 'continue') return cell
      const rec = refreshNestedCellFormulas(cell, r, col, grid)
      if (!rec) return cell
      count += rec.count
      changed = true
      return rec.value
    })
  })
  return changed ? { value: { ...model, rows }, count } : null
}

/**
 * F9 for nested tables: recompute every nested formula against its own table's
 * grid (ABOVE/LEFT resolve inside the nested table, never the outer one) and
 * write the updated models back in one TRACK_IGNORE transaction. Returns the
 * number of recomputed caches (0 = nothing changed).
 */
export function refreshNestedTableFormulas(editor: Editor): number {
  let updated = 0
  const tr = editor.state.tr
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docNestedTable') return
    const model = node.attrs.model as TableModel | null
    if (!model?.rows?.length) return
    const rec = refreshNestedModelFormulas(model)
    if (!rec) return
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, model: rec.value })
    updated += rec.count
  })
  if (updated > 0) editor.view.dispatch(tr.setMeta(TRACK_IGNORE, true))
  return updated
}

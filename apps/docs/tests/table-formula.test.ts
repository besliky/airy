import { TextSelection } from '@tiptap/pm/state'
import { DOMSerializer, Fragment, type Node as PmModelNode } from '@tiptap/pm/model'
import { parseDocx, type Block, type TableModel } from '@airy-office/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  pmTableToModel,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyFieldCaches } from '../src/renderer/editor/revisions'
import {
  collectTableFormulaJobs,
  insertCellFormula,
  refreshNestedTableFormulas,
  proposeCellFormula,
} from '../src/renderer/editor/table-formulas'
import type { Editor } from '@tiptap/core'

const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>10</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>apples</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>20</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>pears</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>30</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p/></w:tc></w:tr></w:tbl>'

/** same shape but with Word's fldSimple form (direct w:p child) in the last cell */
const PARSED_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>10</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>20</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:fldSimple w:instr=" =SUM(ABOVE) \\# &quot;#,##0.00&quot; ">' +
  '<w:r><w:t>30.00</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>'

/**
 * Outer 2×2 (1, 5 / nested, formula) whose (1,0) cell holds a nested 2×2
 * table (7, 4 / 6, =SUM(ABOVE) with a stale '99' cache). `outerInstr` selects
 * the outer formula instruction.
 */
const nestedTableXml = (outerInstr: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>7</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>6</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:fldSimple w:instr="=SUM(ABOVE)"><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>' +
  '</w:tbl>' +
  '<w:p/></w:tc>' +
  `<w:tc><w:p><w:fldSimple w:instr="${outerInstr}"><w:r><w:t>5</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>` +
  '</w:tbl>'

/**
 * Two levels of nesting: the outer table's cell holds table L1 (7, [table L2] /
 * 8, =SUM(LEFT) stale cache '3'); L2 holds 4, 3 / 6, =SUM(ABOVE) stale cache '99'.
 */
const TWO_LEVEL =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>o</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>7</w:t></w:r></w:p></w:tc>' +
  '<w:tc>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="750"/><w:gridCol w:w="750"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>6</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:fldSimple w:instr="=SUM(ABOVE)"><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>' +
  '</w:tbl>' +
  '<w:p/></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>8</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:fldSimple w:instr="=SUM(LEFT)"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>' +
  '</w:tbl>' +
  '<w:p/></w:tc>' +
  '<w:tc><w:p><w:fldSimple w:instr="=SUM(ABOVE)"><w:r><w:t>5</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>' +
  '</w:tbl>'

/** a different table shape for paste targets: 100 / 200 / empty formula slot */
const DEST_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>200</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'

async function openDoc(bodyXml: string) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const editor = createTrackedEditor({
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

/** place the caret inside the cell whose text contains `text` */
function caretInCell(editor: Editor, text: string): void {
  let cellPos = -1
  editor.state.doc.descendants((node, pos) => {
    if (cellPos < 0 && node.type.name === 'docTableCell' && node.textContent.includes(text)) {
      cellPos = pos
    }
  })
  expect(cellPos).toBeGreaterThan(-1)
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(cellPos + 1))),
  )
}

function caretInLastCell(editor: Editor): void {
  let cellPos = -1
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableCell') cellPos = pos
  })
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(cellPos + 1))),
  )
}

/** the (single) text node carrying a tableFormula mark */
function formulaTextOf(editor: Editor): { text: string; instr: string } | null {
  let found: { text: string; instr: string } | null = null
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === 'tableFormula')
    if (mark && !found) found = { text: node.text ?? '', instr: String(mark.attrs.instr) }
  })
  return found
}

/** every text node carrying a tableFormula mark, in document order */
function formulaTexts(editor: Editor): Array<{ text: string; instr: string }> {
  const found: Array<{ text: string; instr: string }> = []
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === 'tableFormula')
    if (mark) found.push({ text: node.text ?? '', instr: String(mark.attrs.instr) })
  })
  return found
}

/** formula runs inside docNestedTable models (recursing into deeper levels) */
function nestedCacheTexts(editor: Editor): Array<{ text: string; instr: string }> {
  const out: Array<{ text: string; instr: string }> = []
  const walkModel = (model: TableModel) =>
    model.rows.forEach((row) =>
      row.forEach((cell) => {
        cell.richParas?.forEach((p) =>
          p.runs.forEach((run) => {
            if (run.formulaField !== undefined) {
              out.push({ text: run.text, instr: String(run.formulaField) })
            }
          }),
        )
        cell.nestedTables?.forEach(walkModel)
      }),
    )
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docNestedTable') walkModel(node.attrs.model as TableModel)
  })
  return out
}

/** the first docNestedTable atom of the document */
function nestedAtomOf(editor: Editor): { pos: number; model: TableModel } {
  const hits: Array<{ pos: number; model: TableModel }> = []
  editor.state.doc.descendants((node, pos) => {
    if (!hits.length && node.type.name === 'docNestedTable') {
      hits.push({ pos, model: node.attrs.model as TableModel })
    }
  })
  expect(hits).toHaveLength(1)
  return hits[0]!
}

/** edit one nested-model cell's text the way the nested-table island commits it
 * (plain paras written back, richParas dropped) */
function setNestedCellText(
  editor: Editor,
  row: number,
  col: number,
  text: string,
  model?: TableModel,
): void {
  const atom = nestedAtomOf(editor)
  const target = model ?? atom.model
  const rows = target.rows.map((r, ri) =>
    ri === row
      ? r.map((c, ci) => {
          if (ci !== col) return c
          const copy = { ...c, paras: [text] }
          delete copy.richParas
          return copy
        })
      : r,
  )
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(atom.pos, undefined, { model: { ...target, rows } }),
  )
}

/** clipboard HTML for the (single) marked run, as the copy side renders it */
function markedRunHtml(editor: Editor): string {
  const marked: PmModelNode[] = []
  editor.state.doc.descendants((node) => {
    if (!marked.length && node.isText && node.marks.some((m) => m.type.name === 'tableFormula')) {
      marked.push(node)
    }
  })
  expect(marked).toHaveLength(1)
  const dom = DOMSerializer.fromSchema(editor.schema).serializeFragment(Fragment.from(marked[0]!))
  const div = document.createElement('div')
  div.appendChild(dom)
  return div.innerHTML
}

function savePlanTableXml(editor: Editor, blocks: Block[]) {
  const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, blocks)
  return plan.saveBlocks
    .map((b) => (b as { xml?: string }).xml ?? '')
    .find((x) => x.includes('w:tbl'))
}

/** replace one cell's text via a plain transaction */
function replaceCellText(editor: Editor, from: string, to: string): void {
  const tr = editor.state.tr
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text === from) tr.insertText(to, pos, pos + node.nodeSize)
  })
  editor.view.dispatch(tr)
}

// tracked editors are destroyed before the jsdom environment goes away
afterEach(() => drainTrackedEditors())

describe('table formula fields', () => {
  it('shows the cached result of a parsed fldSimple and keeps the instruction', async () => {
    const { editor } = await openDoc(PARSED_TABLE)
    expect(formulaTextOf(editor)).toEqual({ text: '30.00', instr: '=SUM(ABOVE) \\# "#,##0.00"' })
    // the save model round-trips the run; Word's own cached text stays put
    const table = editor.getJSON().content![0] as unknown as PmNode
    const run = pmTableToModel(table).rows[2][0].richParas?.[0].runs[0]
    expect(run?.formulaField).toBe('=SUM(ABOVE) \\# "#,##0.00"')
    expect(run?.text).toBe('30.00')
  })

  it('proposes =SUM(ABOVE) like Word and computes the cache on insert', async () => {
    const { editor } = await openDoc(TABLE)
    caretInCell(editor, '30')
    expect(proposeCellFormula(editor)).toBe('=SUM(ABOVE)')
    // the formula replaces the target cell's text; the sum is 10 + 20
    expect(insertCellFormula(editor, '=SUM(ABOVE)')).toBe(true)
    expect(formulaTextOf(editor)).toEqual({ text: '30', instr: '=SUM(ABOVE)' })
  })

  it('rejects a non-formula instruction and a caret outside any table', async () => {
    const { editor } = await openDoc(TABLE)
    caretInCell(editor, '30')
    expect(insertCellFormula(editor, 'SUM(ABOVE)')).toBe(false)
    const standalone = createTrackedEditor({
      extensions: editorExtensions,
      content: blocksToPmDoc([
        {
          id: 'p1',
          type: 'passthrough',
          docxIndex: null,
          originalXml: '<w:p/>',
          runs: [{ text: 'plain' }],
        } as unknown as Block,
      ]),
    })
    caretInLastCell(standalone)
    expect(insertCellFormula(standalone, '=SUM(ABOVE)')).toBe(false)
  })

  it('applies the numeric picture switch to the cached result on insert', async () => {
    const { editor } = await openDoc(TABLE)
    caretInCell(editor, '30')
    insertCellFormula(editor, '=SUM(ABOVE) \\# "#,##0.00"')
    expect(formulaTextOf(editor)?.text).toBe('30.00')
  })

  it('caches Undefined Bookmark when the scanned direction is empty', async () => {
    const { editor } = await openDoc(TABLE)
    // the last cell (row 3, column B): the column above holds only text cells
    caretInLastCell(editor)
    expect(insertCellFormula(editor, '=SUM(ABOVE)')).toBe(true)
    expect(formulaTextOf(editor)?.text).toBe('Undefined Bookmark')
  })

  it('F9 recomputes after the referenced cells change', async () => {
    const { editor } = await openDoc(TABLE)
    caretInCell(editor, '30')
    insertCellFormula(editor, '=SUM(ABOVE)')
    expect(formulaTextOf(editor)?.text).toBe('30')
    replaceCellText(editor, '10', '12')
    const jobs = collectTableFormulaJobs(editor)
    expect(jobs).toHaveLength(1)
    expect(jobs[0].text).toBe('32')
    applyFieldCaches(editor, jobs)
    expect(formulaTextOf(editor)?.text).toBe('32')
  })

  it('F9 turns the cache into Undefined Bookmark when the column empties', async () => {
    const { editor } = await openDoc(TABLE)
    caretInCell(editor, '30')
    insertCellFormula(editor, '=SUM(ABOVE)')
    replaceCellText(editor, '10', '')
    replaceCellText(editor, '20', '')
    const jobs = collectTableFormulaJobs(editor)
    expect(jobs).toHaveLength(1)
    applyFieldCaches(editor, jobs)
    expect(formulaTextOf(editor)?.text).toBe('Undefined Bookmark')
  })

  it('saves the cell as w:fldSimple with instruction and cached result', async () => {
    const { editor, parsed } = await openDoc(TABLE)
    caretInCell(editor, '30')
    insertCellFormula(editor, '=SUM(ABOVE)')
    const xml = savePlanTableXml(editor, parsed.blocks)
    expect(xml).toContain('<w:fldSimple w:instr="=SUM(ABOVE)">')
    expect(xml).toContain('>30</w:t>')
  })

  it('keeps the numeric picture switch and F9 result on save of a parsed formula', async () => {
    const { editor, parsed } = await openDoc(PARSED_TABLE)
    replaceCellText(editor, '10', '12')
    const jobs = collectTableFormulaJobs(editor)
    expect(jobs).toHaveLength(1)
    expect(jobs[0].text).toBe('32.00')
    applyFieldCaches(editor, jobs)
    const xml = savePlanTableXml(editor, parsed.blocks)
    // attribute values escape the quotes of the \# picture
    expect(xml).toContain('w:instr="=SUM(ABOVE) \\# &quot;#,##0.00&quot;"')
    expect(xml).toContain('>32.00</w:t>')
  })

  it('recomputes nested-table formulas on their own grid (BUG-1757)', async () => {
    const { editor } = await openDoc(nestedTableXml('=SUM(ABOVE)'))
    // edit the inner source 4 → 10 the way the nested-table island commits it
    setNestedCellText(editor, 0, 1, '10')
    // the outer =SUM(ABOVE) still reads 5 (the cell above it) — no PM job, and
    // it never takes the edited inner value
    expect(collectTableFormulaJobs(editor)).toEqual([])
    // one F9: the stale inner cache 99 recomputes by the INNER grid (= 10)
    expect(refreshNestedTableFormulas(editor)).toBe(1)
    expect(nestedCacheTexts(editor)).toEqual([{ text: '10', instr: '=SUM(ABOVE)' }])
    // F9 is idempotent once the caches are fresh
    expect(refreshNestedTableFormulas(editor)).toBe(0)
  })

  it('outer formulas never read the nested cells as one number (BUG-1757)', async () => {
    const { editor } = await openDoc(nestedTableXml('=SUM(LEFT)'))
    setNestedCellText(editor, 0, 1, '10')
    // the outer LEFT neighbor holds a table, not a number: Word stops the scan
    // there (Undefined Bookmark) instead of summing the nested cells' text
    const jobs = collectTableFormulaJobs(editor)
    expect(jobs).toEqual([expect.objectContaining({ text: 'Undefined Bookmark' })])
    applyFieldCaches(editor, jobs)
    expect(formulaTexts(editor).map((t) => t.text)).toEqual(['Undefined Bookmark'])
    // the outer scan never touches the nested table's own caches
    expect(nestedCacheTexts(editor)).toEqual([{ text: '99', instr: '=SUM(ABOVE)' }])
  })

  it('recomputes each nested level on its own grid (BUG-1757, two levels)', async () => {
    const { editor } = await openDoc(TWO_LEVEL)
    // edit L2's source 3 → 10 through the model (as the deepest island commits)
    const atom = nestedAtomOf(editor)
    const l2 = atom.model.rows[0]![1]!.nestedTables![0]!
    const l2Rows = l2.rows.map((r, ri) =>
      ri === 0
        ? r.map((c, ci) => {
            if (ci !== 1) return c
            const copy = { ...c, paras: ['10'] }
            delete copy.richParas
            return copy
          })
        : r,
    )
    const rows = atom.model.rows.map((r, ri) => {
      if (ri !== 0) return r
      return r.map((c, ci) => (ci === 1 ? { ...c, nestedTables: [{ ...l2, rows: l2Rows }] } : c))
    })
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(atom.pos, undefined, { model: { ...atom.model, rows } }),
    )
    // one F9 recomputes both levels, each by its own grid: L2's =SUM(ABOVE)
    // follows its own column (10), L1's =SUM(LEFT) sees only its own row (8) —
    // never the other level's cells
    expect(refreshNestedTableFormulas(editor)).toBe(2)
    expect(nestedCacheTexts(editor)).toEqual([
      { text: '10', instr: '=SUM(ABOVE)' }, // L2: its own column 3 → 10
      { text: '8', instr: '=SUM(LEFT)' }, // L1: its own row neighbor 8
    ])
    // the outer =SUM(ABOVE) cache 5 is untouched by the nested levels
    expect(formulaTexts(editor).map((t) => t.text)).toEqual(['5'])
  })

  it('saves recomputed caches for the nested level and the outer level (BUG-1757)', async () => {
    const { editor, parsed } = await openDoc(nestedTableXml('=SUM(ABOVE)'))
    setNestedCellText(editor, 0, 1, '10')
    refreshNestedTableFormulas(editor)
    const xml = savePlanTableXml(editor, parsed.blocks)
    expect((xml?.match(/<w:fldSimple w:instr="=SUM\(ABOVE\)">/g) ?? []).length).toBe(2)
    expect(xml).toContain('>10</w:t>') // inner cache recomputed
    expect(xml).not.toContain('99') // the stale inner cache is gone
  })

  it('paste between tables keeps the instruction and F9 recomputes at the new position (BUG-1758)', async () => {
    const source = await openDoc(TABLE)
    caretInCell(source.editor, '30')
    insertCellFormula(source.editor, '=SUM(ABOVE)') // cache 30 in the source grid
    const html = markedRunHtml(source.editor)
    // paste into a cell of a DIFFERENT table whose column holds 100 + 200
    const dest = await openDoc(DEST_TABLE)
    caretInLastCell(dest.editor)
    dest.editor.commands.insertContent(html)
    // Word shows the copied cache until the field updates…
    expect(formulaTextOf(dest.editor)).toEqual({ text: '30', instr: '=SUM(ABOVE)' })
    // …then one F9 resolves the same instruction against the destination grid
    const jobs = collectTableFormulaJobs(dest.editor)
    expect(jobs).toEqual([expect.objectContaining({ text: '300' })])
    applyFieldCaches(dest.editor, jobs)
    expect(formulaTextOf(dest.editor)).toEqual({ text: '300', instr: '=SUM(ABOVE)' })
  })

  it('does not F9-write or save a degenerate empty-instruction field (BUG-1758)', async () => {
    const { editor, parsed } = await openDoc(TABLE)
    caretInLastCell(editor)
    editor.commands.insertContent('<span data-table-formula="">300</span>')
    // F9 leaves the cached text alone (no !Syntax Error rewrite)
    expect(collectTableFormulaJobs(editor)).toEqual([])
    const xml = savePlanTableXml(editor, parsed.blocks)
    expect(xml).not.toContain('fldSimple')
    expect(xml).toContain('>300</w:t>')
  })
})

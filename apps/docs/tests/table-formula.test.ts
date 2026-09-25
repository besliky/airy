import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, type Block } from '@airy-office/docx-engine'
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
})

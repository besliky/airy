import { undoDepth } from '@tiptap/pm/history'
import { parseDocx, type TableModel } from '@airy-office/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyFieldCaches } from '../src/renderer/editor/revisions'
import {
  collectTableFormulaJobs,
  refreshNestedTableFormulas,
} from '../src/renderer/editor/table-formulas'
import type { Editor } from '@tiptap/core'

/**
 * UX-1763: F9 recomputes field caches, so a refresh is recomputation, not an
 * authored edit — the transactions that write caches (applyFieldCaches,
 * refreshNestedTableFormulas) must stay out of the prosemirror-history undo
 * stack. Otherwise every F9 press eats one Ctrl+Z step and F9 spam buries the
 * user's real edits (the docs twin of the sheets finding fixed in #261).
 *
 * Grouping note: prosemirror-history merges a recorded transaction into the
 * previous event only when its changed ranges overlap the previous event's
 * ranges within newGroupDelay. The churn test alternates edits between two
 * source cells (and the cache write always lands in a third cell), so every
 * authored edit and every pre-fix cache write forms its own undo event and the
 * undoDepth arithmetic below is exact.
 */

const PARSED_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>10</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>20</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:fldSimple w:instr="=SUM(ABOVE) \\# &quot;#,##0.00&quot; ">' +
  '<w:r><w:t>30.00</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl>'

/** outer table whose (1,0) cell holds a nested 2×2 (7, 4 / 6, =SUM(ABOVE) with a stale '99' cache) */
const NESTED_TABLE =
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
  '<w:tc><w:p><w:fldSimple w:instr="=SUM(ABOVE)"><w:r><w:t>5</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>' +
  '</w:tbl>'

async function openDoc(bodyXml: string): Promise<Editor> {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  return createTrackedEditor({
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

/** replace one cell's text via a plain (undoable) transaction */
function replaceCellText(editor: Editor, from: string, to: string): void {
  const tr = editor.state.tr
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text === from) tr.insertText(to, pos, pos + node.nodeSize)
  })
  editor.view.dispatch(tr)
}

const hasText = (editor: Editor, text: string): boolean => {
  let found = false
  editor.state.doc.descendants((node) => {
    if (node.isText && node.text === text) found = true
  })
  return found
}

/** the (single) text node carrying a tableFormula mark */
function formulaCacheText(editor: Editor): string | null {
  let found: string | null = null
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === 'tableFormula')
    if (mark && !found) found = node.text ?? ''
  })
  return found
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

/** visible text of one nested-model cell (the way the island renders it) */
function nestedCellText(editor: Editor, row: number, col: number): string {
  const cell = nestedAtomOf(editor).model.rows[row]![col]!
  return cell.richParas?.length
    ? cell.richParas.map((p) => p.runs.map((run) => run.text).join('')).join('\n')
    : (cell.paras ?? []).join('\n')
}

/** edit one nested-model cell's text the way the nested-table island commits it */
function setNestedCellText(editor: Editor, row: number, col: number, text: string): void {
  const atom = nestedAtomOf(editor)
  const rows = atom.model.rows.map((r, ri) =>
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
    editor.state.tr.setNodeMarkup(atom.pos, undefined, { model: { ...atom.model, rows } }),
  )
}

// tracked editors are destroyed before the jsdom environment goes away
afterEach(() => drainTrackedEditors())

describe('F9 cache writes stay out of the undo history (UX-1763)', () => {
  it('one Ctrl+Z reverts the edit after an F9 spam over a static document', async () => {
    const editor = await openDoc(PARSED_TABLE)
    // the user's authored edit: source 10 → 12
    replaceCellText(editor, '10', '12')
    const depthAfterEdit = undoDepth(editor.state)
    expect(depthAfterEdit).toBeGreaterThan(0)
    // F9 ×50: the first press rewrites the stale cache, the rest find nothing
    // to recompute and must not dispatch (or record) anything at all
    for (let i = 0; i < 50; i++) applyFieldCaches(editor, collectTableFormulaJobs(editor))
    expect(formulaCacheText(editor)).toBe('32.00') // caches applied
    expect(undoDepth(editor.state)).toBe(depthAfterEdit) // history not grown
    // the FM-5 repro: a single Ctrl+Z must land on the user's edit…
    expect(editor.commands.undo()).toBe(true)
    expect(hasText(editor, '10')).toBe(true)
    expect(hasText(editor, '12')).toBe(false)
    // …and the field still recomputes from the reverted content afterwards
    const jobs = collectTableFormulaJobs(editor)
    applyFieldCaches(editor, jobs)
    expect(formulaCacheText(editor)).toBe('30.00')
  })

  it('50 edit+F9 pairs add exactly 50 undo events — the cache writes add none', async () => {
    const editor = await openDoc(PARSED_TABLE)
    const depthStart = undoDepth(editor.state)
    // churn: 50 (edit + F9) pairs. The edit alternates between the two source
    // cells on a 4-step cycle, so each edit forms its own undo event (every
    // edit's range differs from the previous event's), and every press
    // recomputes a real cache change in the formula cell (a third position)
    const cycle = [
      ['10', '11'],
      ['20', '21'],
      ['11', '10'],
      ['21', '20'],
    ] as const
    for (let i = 0; i < 50; i++) {
      const [from, to] = cycle[i % 4]
      replaceCellText(editor, from, to)
      const jobs = collectTableFormulaJobs(editor)
      expect(jobs).toHaveLength(1) // every press really rewrites the cache
      applyFieldCaches(editor, jobs)
    }
    expect(formulaCacheText(editor)).toBe('32.00') // 11 + 21, caches applied
    expect(undoDepth(editor.state)).toBe(depthStart + 50)
    // one undo steps onto the last authored edit (21 → 20), not an F9 record
    expect(editor.commands.undo()).toBe(true)
    expect(hasText(editor, '20')).toBe(true)
    expect(hasText(editor, '21')).toBe(false)
    expect(formulaCacheText(editor)).toBe('32.00') // mute write survived the undo
    // the next F9 recomputes from the reverted sources
    applyFieldCaches(editor, collectTableFormulaJobs(editor))
    expect(formulaCacheText(editor)).toBe('31.00')
  })

  it('nested-table F9 refresh is mute: the refresh adds no undo event', async () => {
    const editor = await openDoc(NESTED_TABLE)
    // authored edit through the island's commit path: inner source 4 → 10
    setNestedCellText(editor, 0, 1, '10')
    const depthAfterEdit = undoDepth(editor.state)
    // F9 for nested models: rewrites the stale '99' cache to '10' in one
    // transaction — recomputation, so it must not add an undo event
    expect(refreshNestedTableFormulas(editor)).toBe(1)
    expect(nestedCellText(editor, 1, 1)).toBe('10') // cache applied
    expect(undoDepth(editor.state)).toBe(depthAfterEdit)
    // a single undo steps onto the authored edit. The nested cache lives
    // inside the same `model` attribute the edit rewrites, so it rolls back
    // with the content (the sheets-side UX-1763 semantics) — the stale
    // pre-edit cache shows until the next F9 recomputes it
    expect(editor.commands.undo()).toBe(true)
    expect(nestedCellText(editor, 0, 1)).toBe('4') // the edit is reverted
    expect(nestedCellText(editor, 1, 1)).toBe('99') // …together with its cache
    expect(refreshNestedTableFormulas(editor)).toBe(1) // and F9 still works
    expect(nestedCellText(editor, 1, 1)).toBe('4')
  })
})

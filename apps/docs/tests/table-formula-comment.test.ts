import { parseDocx, type Block } from '@airy-office/docx-engine'
import type { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  type PmNode as PmJsonNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyFieldCaches } from '../src/renderer/editor/revisions'
import { collectTableFormulaJobs } from '../src/renderer/editor/table-formulas'

const COMMENTS_PART = {
  path: 'word/comments.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:comment w:id="0" w:author="Alice" w:date="2026-01-01T00:00:00Z">' +
    '<w:p><w:r><w:t>check this total</w:t></w:r></w:p></w:comment></w:comments>',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
}

const COMMENTS_REL =
  '<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'

/**
 * BUG-1756 shape (audit f_comment.docx): the =SUM(ABOVE) complex field carries
 * no cached result (separate → end with nothing between) and the displayed
 * value follows the field as a plain run — the whole cell content wrapped in a
 * comment range.
 */
const COMMENT_FORMULA_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>250</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>350</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p>' +
  '<w:commentRangeStart w:id="0"/>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> =SUM(ABOVE) </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t>600</w:t></w:r>' +
  '<w:commentRangeEnd w:id="0"/>' +
  '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>' +
  '</w:p></w:tc></w:tr></w:tbl>'

async function openDoc(bodyXml: string) {
  const source = await buildDocx({
    bodyXml,
    extraParts: [COMMENTS_PART],
    extraRels: COMMENTS_REL,
  })
  const parsed = await parseDocx(source)
  const editor = createTrackedEditor({
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

interface FormulaRunJson {
  text: string
  marks: string[]
}

function formulaRuns(editor: Editor): FormulaRunJson[] {
  const found: FormulaRunJson[] = []
  editor.state.doc.descendants((node) => {
    if (node.isText && node.marks.some((m) => m.type.name === 'tableFormula')) {
      found.push({ text: node.text ?? '', marks: node.marks.map((m) => m.type.name) })
    }
  })
  return found
}

function cellTexts(editor: Editor): string[] {
  const texts: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docTableCell') texts.push(node.textContent)
  })
  return texts
}

/** replace one cell's text via a plain transaction */
function replaceCellText(editor: Editor, from: string, to: string) {
  const tr = editor.state.tr
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text === from) tr.insertText(to, pos, pos + node.nodeSize)
  })
  editor.view.dispatch(tr)
}

function savePlanTableXml(editor: Editor, blocks: Block[]) {
  const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmJsonNode, blocks)
  return plan.saveBlocks
    .map((b) => (b as { xml?: string }).xml ?? '')
    .find((x) => x.includes('w:tbl'))
}

// tracked editors are destroyed before the jsdom environment goes away
afterEach(() => drainTrackedEditors())

describe('table formula under a comment range (BUG-1756)', () => {
  it('opens the empty-cache field with the value run as its single formula cache', async () => {
    const { editor } = await openDoc(COMMENT_FORMULA_TABLE)
    // the plain "600" after the field folds into the formula run: one cache,
    // no placeholder space, the comment mark rides along
    expect(formulaRuns(editor)).toEqual([{ text: '600', marks: ['tableFormula', 'comment'] }])
    expect(cellTexts(editor)[2]).toBe('600')
  })

  it('keeps a clean F9 a no-op and updates the cache exactly once after an edit', async () => {
    const { editor } = await openDoc(COMMENT_FORMULA_TABLE)
    // pure F9 with no source edits: the cache is already correct
    expect(collectTableFormulaJobs(editor)).toEqual([])
    expect(formulaRuns(editor)).toEqual([{ text: '600', marks: ['tableFormula', 'comment'] }])
    // edit a source cell, then F9: the cache is rewritten in place — one run,
    // one cache, the comment mark survives (audit vector 600 -> 990)
    replaceCellText(editor, '250', '640')
    const jobs = collectTableFormulaJobs(editor)
    expect(jobs).toHaveLength(1)
    expect(jobs[0].text).toBe('990')
    applyFieldCaches(editor, jobs)
    expect(formulaRuns(editor)).toEqual([{ text: '990', marks: ['tableFormula', 'comment'] }])
    expect(cellTexts(editor)[2]).toBe('990')
  })

  it('saves one fldSimple inside the comment range with no stray cache run', async () => {
    const { editor, parsed } = await openDoc(COMMENT_FORMULA_TABLE)
    replaceCellText(editor, '250', '640')
    applyFieldCaches(editor, collectTableFormulaJobs(editor))
    const xml = savePlanTableXml(editor, parsed.blocks) ?? ''
    expect(xml.match(/<w:fldSimple /g)).toHaveLength(1)
    expect(xml).toContain(
      '<w:fldSimple w:instr="=SUM(ABOVE)"><w:r><w:t xml:space="preserve">990</w:t></w:r></w:fldSimple>',
    )
    // exactly one cached value, and the comment range still wraps the field
    expect(xml.match(/>990<\/w:t>/g)).toHaveLength(1)
    expect(xml.indexOf('<w:commentRangeStart w:id="0"')).toBeGreaterThan(-1)
    expect(xml.indexOf('<w:commentRangeEnd w:id="0"')).toBeGreaterThan(
      xml.indexOf('<w:fldSimple w:instr="=SUM(ABOVE)"'),
    )
  })
})

import { parseDocx, saveDocx, type Block, type ParsedDocFull } from '@airy-office/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  isLooseTableRowBlock,
  mergeLooseTableRowBlocks,
} from '../src/renderer/editor/loose-table-rows'
import { insertCellFormula } from '../src/renderer/editor/table-formulas'
import { TextSelection } from '@tiptap/pm/state'
import type { Editor } from '@tiptap/core'

/**
 * BUG-1707: a document whose table rows sit loose at body level (the w:tbl
 * wrapper lost to a broken producer) used to render every row as an opaque
 * raw chip — the whole table invisible and uneditable, silently. The rows now
 * merge into a native editable table, formulas included; a save repairs the
 * body to a valid <w:tbl>. Valid w:tbl tables with any number of table
 * formula fields keep rendering natively (the PAR-103 path).
 */

/** rows x 4 cols of body XML; the first `formulas` cells of the last row carry fldSimple =SUM(LEFT) */
function tableRows(rows: number, formulas: number): string {
  const trs: string[] = []
  for (let r = 0; r < rows; r++) {
    let row = '<w:tr>'
    for (let c = 0; c < 4; c++) {
      if (r === rows - 1 && c < formulas) {
        row += `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:fldSimple w:instr="=SUM(LEFT)"><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p></w:tc>`
      } else {
        row += `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${r}-${c}</w:t></w:r></w:p></w:tc>`
      }
    }
    trs.push(row + '</w:tr>')
  }
  return trs.join('')
}

/** malformed shape this recovery targets: rows sit loose at body level */
function looseRowTable(rows: number, formulas: number): string {
  return (
    '<w:p><w:r><w:t>intro</w:t></w:r></w:p>' +
    tableRows(rows, formulas) +
    '<w:p><w:r><w:t>tail</w:t></w:r></w:p>'
  )
}

/** valid Word-like shape: the same table properly wrapped in w:tbl */
function wrappedTable(rows: number, formulas: number): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
    tableRows(rows, formulas) +
    '</w:tbl><w:p><w:r><w:t>tail</w:t></w:r></w:p>'
  )
}

async function open(bodyXml: string): Promise<{ editor: Editor; parsed: ParsedDocFull }> {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const editor = createTrackedEditor({
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

function snapshot(editor: Editor, parsed: ParsedDocFull) {
  let docTable = 0
  let docProtected = 0
  let chips = 0
  const formulaInstrs: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docTable') docTable++
    if (node.type.name === 'docProtected') {
      docProtected++
      if (String(node.attrs?.label ?? '') === 'w:tr') chips++
    }
    const mark = node.marks?.find((m) => m.type.name === 'tableFormula')
    if (mark) formulaInstrs.push(String(mark.attrs.instr))
  })
  const dom = editor.view.dom as HTMLElement
  return {
    looseRowBlocks: parsed.blocks.filter(isLooseTableRowBlock).length,
    docTable,
    docProtected,
    chips,
    formulaInstrs,
    tds: dom.querySelectorAll('td').length,
    fieldAffordances: dom.querySelectorAll('.doc-ref-field').length,
    showsCellText: dom.textContent?.includes('0-2') && dom.textContent.includes('1-3'),
  }
}

afterEach(() => drainTrackedEditors())

describe('loose body-level table rows (BUG-1707)', () => {
  for (const n of [1, 2, 3]) {
    it(`recovers a malformed table with ${n} formula cell(s) into a native table`, async () => {
      const { editor, parsed } = await open(looseRowTable(3, n))
      const snap = snapshot(editor, parsed)
      expect(snap.looseRowBlocks).toBe(3)
      expect(snap.docTable).toBe(1)
      expect(snap.chips).toBe(0)
      expect(snap.docProtected).toBe(0)
      expect(snap.formulaInstrs).toHaveLength(n)
      expect(snap.formulaInstrs.every((i) => i === '=SUM(LEFT)')).toBe(true)
      // the table's content is visible in the DOM, not chips
      expect(snap.tds).toBe(12)
      expect(snap.fieldAffordances).toBe(n)
      expect(snap.showsCellText).toBe(true)
    })
  }

  it('keeps paragraphs around the recovered table in place', async () => {
    const { editor } = await open(looseRowTable(3, 2))
    const types: string[] = []
    editor.state.doc.forEach((n) => types.push(n.type.name))
    expect(types).toEqual(['docParagraph', 'docTable', 'docParagraph'])
    expect(editor.state.doc.textContent).toContain('intro')
    expect(editor.state.doc.textContent).toContain('tail')
  })

  it('edits the recovered table: inserting a formula into a cell works', async () => {
    // numeric first column so =SUM(ABOVE) has something to add (Word parity)
    const bodyXml =
      '<w:p><w:r><w:t>intro</w:t></w:r></w:p>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>10</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>20</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>'
    const { editor } = await open(bodyXml)
    let cellPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (cellPos < 0 && node.type.name === 'docTableCell' && node.textContent === '') cellPos = pos
    })
    expect(cellPos).toBeGreaterThan(-1)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(cellPos + 1))),
    )
    expect(insertCellFormula(editor, '=SUM(ABOVE)')).toBe(true)
    const texts: string[] = []
    editor.state.doc.descendants((node) => {
      const mark = node.marks?.find((m) => m.type.name === 'tableFormula')
      if (mark) texts.push(node.text ?? '')
    })
    expect(texts).toEqual(['30'])
  })

  it('saves the recovered table as one valid w:tbl with the formulas intact', async () => {
    const bodyXml = looseRowTable(3, 2)
    const source = await buildDocx({ bodyXml })
    const parsed = await parseDocx(source)
    const editor = createTrackedEditor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const tableXmls = plan.saveBlocks
      .map((b) => (b as { xml?: string }).xml ?? '')
      .filter((x) => x.includes('<w:tbl>'))
    // one regenerated table replaces every loose row; no raw rows pass through
    expect(tableXmls).toHaveLength(1)
    expect(tableXmls[0].match(/<w:fldSimple w:instr="=SUM\(LEFT\)">/g) ?? []).toHaveLength(2)
    expect(tableXmls[0]).toContain('<w:tblGrid>')
    // round trip: the repaired file reopens natively with both formulas
    const out = await saveDocx(parsed, plan.saveBlocks, {})
    const reparsed = await parseDocx(new Uint8Array(out))
    const editor2 = createTrackedEditor({
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks) as never,
    })
    const snap = snapshot(editor2, reparsed)
    expect(snap.docTable).toBe(1)
    expect(snap.chips).toBe(0)
    expect(snap.formulaInstrs).toEqual(['=SUM(LEFT)', '=SUM(LEFT)'])
  })

  it('valid wrapped tables with 2-3 formulas render natively (PAR-103 path intact)', async () => {
    for (const n of [2, 3]) {
      const { editor, parsed } = await open(wrappedTable(3, n))
      const snap = snapshot(editor, parsed)
      expect(snap.looseRowBlocks).toBe(0)
      expect(snap.docTable).toBe(1)
      expect(snap.docProtected).toBe(0)
      expect(snap.formulaInstrs).toHaveLength(n)
      expect(snap.tds).toBe(12)
      expect(snap.fieldAffordances).toBe(n)
      expect(snap.showsCellText).toBe(true)
    }
  })

  it('merge leaves documents without loose rows untouched', () => {
    const para: Block = {
      id: 'b0',
      type: 'paragraph',
      docxIndex: 0,
      originalXml: '<w:p/>',
      runs: [{ text: 'x' }],
    }
    const other: Block = {
      id: 'b1',
      type: 'passthrough',
      docxIndex: 1,
      originalXml: '<w:custom/>',
      label: 'w:custom',
    }
    const blocks = [para, other]
    expect(mergeLooseTableRowBlocks(blocks)).toBe(blocks)
  })

  it('merge groups only consecutive loose rows and keeps interleaved paragraphs', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>' +
        '<w:p><w:r><w:t>mid</w:t></w:r></w:p>' +
        '<w:tr><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>',
    })
    const parsed = await parseDocx(source)
    const merged = mergeLooseTableRowBlocks(parsed.blocks).filter((b) => !b.hidden)
    const kinds = merged.map((b) => b.type)
    // row, paragraph, row → two separate single-row tables around the paragraph
    expect(kinds).toEqual(['table', 'paragraph', 'table'])
    expect(merged.filter((b) => b.type === 'table')).toHaveLength(2)
  })
})

import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type ParsedDocFull } from '@airy-office/docx-engine'
import { beforeAll, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { loadLocale } from '../src/renderer/i18n/locale'
import {
  parseSortDate,
  parseSortNumber,
  sortScope,
  sortSelectedParagraphs,
  sortTableRows,
  type SortLevel,
  type SortOptions,
} from '../src/renderer/editor/sort'

// PERF-904: dictionaries load per locale; the zh table these assertions
// match must be loaded first, mirroring the bootstrap-time load in main.tsx.
beforeAll(() => loadLocale('zh'))

/* ================= Word-style key parsing ================= */

describe('sort key parsing', () => {
  it('reads numbers with separators, currency, and accounting negatives', () => {
    expect(parseSortNumber('42')).toBe(42)
    expect(parseSortNumber(' 1,300 ')).toBe(1300)
    expect(parseSortNumber('$1,234.50')).toBe(1234.5)
    expect(parseSortNumber('(45)')).toBe(-45)
    expect(parseSortNumber('3,14')).toBe(3.14)
    expect(parseSortNumber('1.234,56')).toBe(1234.56)
    expect(parseSortNumber('-7.5')).toBe(-7.5)
    expect(parseSortNumber('50%')).toBe(50)
    expect(parseSortNumber('n/a')).toBeNull()
    expect(parseSortNumber('')).toBeNull()
  })

  it('treats non-Word numerals as text, not numbers (BUG-707)', () => {
    // Number() alone accepts these; Word reads them as text -> sort last
    expect(parseSortNumber('0x10')).toBeNull() // hex literal
    expect(parseSortNumber('1e3')).toBeNull() // scientific
    expect(parseSortNumber('Infinity')).toBeNull()
    expect(parseSortNumber('1 2')).toBeNull() // digits glued around a bare space
    // grouped spacing is still a thousands separator, not gluing
    expect(parseSortNumber('1 300')).toBe(1300)
    expect(parseSortNumber('1 300 000')).toBe(1300000)
    expect(parseSortNumber('50 %')).toBe(50)
  })

  it('reads common date formats', () => {
    expect(parseSortDate('2026-03-14')).not.toBeNull()
    expect(parseSortDate('25/12/2026')).not.toBeNull() // day forced by >12
    expect(parseSortDate('03/25/2026')).not.toBeNull() // month forced by >12
    expect(parseSortDate('14 Mar 2026')).not.toBeNull()
    expect(parseSortDate('Mar 14, 2026')).not.toBeNull()
    expect(parseSortDate('2026年3月14日')).not.toBeNull()
    expect(parseSortDate('2026-03-14T08:30')).not.toBeNull()
    expect(parseSortDate('hello')).toBeNull()
    expect(parseSortDate('2026-13-01')).toBeNull() // invalid month
  })

  it('orders numeric and date keys sensibly', () => {
    expect(parseSortDate('2026-01-05')!).toBeGreaterThan(parseSortDate('2025-12-01')!)
    expect(parseSortDate('31/12/99')!).toBeLessThan(parseSortDate('01/01/26')!)
  })

  it('rejects rollover dates instead of normalizing them (BUG-705)', () => {
    // Date.UTC would silently roll these into the next month/day; Word
    // treats them as text, so they must read as null (sort last)
    expect(parseSortDate('30/02/2026')).toBeNull() // Feb 30 -> Mar 2
    expect(parseSortDate('31/04/2026')).toBeNull() // Apr 31 -> May 1
    expect(parseSortDate('29/02/2023')).toBeNull() // non-leap Feb 29
    expect(parseSortDate('2026-03-14T25:00')).toBeNull() // hour 25 rolls a day
    expect(parseSortDate('2026-03-14T23:59')).not.toBeNull()
    // leap years keep Feb 29
    expect(parseSortDate('29/02/2024')).not.toBeNull()
  })

  it('reads short ISO years literally, not as 1900+year (BUG-705)', () => {
    // Date.UTC maps years 0-99 to 1900+y; the written year must win
    expect(new Date(parseSortDate('0066-05-05')!).getUTCFullYear()).toBe(66)
    expect(parseSortDate('0066-05-05')!).toBeLessThan(parseSortDate('1000-01-01')!)
  })

  it('applies the Office 2029 pivot to two-digit years (BUG-706)', () => {
    // 00-29 -> 2000s, 30-99 -> 1900s (the split Excel 97 used, 00-39, is
    // not what current Word/Excel apply)
    expect(new Date(parseSortDate('01/01/29')!).getUTCFullYear()).toBe(2029)
    expect(new Date(parseSortDate('31/12/30')!).getUTCFullYear()).toBe(1930)
    expect(new Date(parseSortDate('15/06/45')!).getUTCFullYear()).toBe(1945)
    expect(parseSortDate('31/12/30')!).toBeLessThan(parseSortDate('01/01/29')!)
  })
})

/* ================= table sorting ================= */

const HEAD = ['Name', 'City', 'Qty', 'Date', 'Note']
const DATA = [
  ['Delta', 'ny', '1300', '2026-03-14', 'd4'],
  ['Alpha', 'la', '25', '2025-12-01', 'a1'],
  ['Charlie', 'la', '3', '2026-01-05', 'c3'],
  ['Bravo', 'ny', '470', '2024-07-21', 'b2'],
]

const tc = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const tr = (cells: string[], header = false) =>
  `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${cells.map(tc).join('')}</w:tr>`
const GRID5 = '<w:tblGrid>' + '<w:gridCol w:w="1800"/>'.repeat(5) + '</w:tblGrid>'
const TABLE5 = `<w:tbl><w:tblPr/>${GRID5}${tr(HEAD, true)}${DATA.map((r) => tr(r)).join('')}</w:tbl>`

async function openTable(xml: string): Promise<{ editor: Editor; parsed: ParsedDocFull }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: xml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let firstCell = -1
  editor.state.doc.descendants((node, pos) => {
    if (
      firstCell < 0 &&
      (node.type.name === 'docTableCell' || node.type.name === 'docTableHeader')
    ) {
      firstCell = pos
    }
  })
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(firstCell + 1))),
  )
  return { editor, parsed }
}

/** first-column texts of the table rows, in document order */
function rowNames(editor: Editor): string[] {
  const table = editor.state.doc.firstChild!
  const names: string[] = []
  table.forEach((row) => {
    names.push(row.firstChild?.textContent ?? '')
  })
  return names
}

const level = (column: number, type: SortLevel['type'], descending = false): SortLevel => ({
  column,
  type,
  descending,
})
const options = (levels: SortLevel[], headerRow = true): SortOptions => ({ levels, headerRow })

/** sorted doc bytes: PM sort -> save plan -> saveDocx */
async function resaved(editor: Editor, parsed: ParsedDocFull): Promise<Uint8Array> {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  return saveDocx(parsed, plan.saveBlocks)
}

describe('table sorting', () => {
  it('sorts rows by text ascending and descending, keeping the header row', async () => {
    const { editor, parsed } = await openTable(TABLE5)
    const scope = sortScope(editor.state)
    expect(scope).toMatchObject({ kind: 'table', columnCount: 5, rowCount: 5, headerRow: true })
    expect((scope as { headerLabels: string[] }).headerLabels).toEqual(HEAD)

    expect(sortTableRows(options([level(0, 'text')]))(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    expect(rowNames(editor)).toEqual(['Name', 'Alpha', 'Bravo', 'Charlie', 'Delta'])
    expect(
      (await parseDocx(await resaved(editor, parsed))).blocks[0].table!.rows.map((r) =>
        r[0].paras.join(''),
      ),
    ).toEqual(['Name', 'Alpha', 'Bravo', 'Charlie', 'Delta'])
    editor.destroy()
  })

  it('sorts descending', async () => {
    const { editor } = await openTable(TABLE5)
    expect(
      sortTableRows(options([level(0, 'text', true)]))(editor.state, editor.view.dispatch),
    ).toBe(true)
    expect(rowNames(editor)).toEqual(['Name', 'Delta', 'Charlie', 'Bravo', 'Alpha'])
    editor.destroy()
  })

  it('sorts numerically (thousands separators read as numbers)', async () => {
    const { editor, parsed } = await openTable(TABLE5)
    expect(sortTableRows(options([level(2, 'number')]))(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    expect(rowNames(editor)).toEqual(['Name', 'Charlie', 'Alpha', 'Bravo', 'Delta'])
    expect(
      (await parseDocx(await resaved(editor, parsed))).blocks[0].table!.rows.map((r) =>
        r[0].paras.join(''),
      ),
    ).toEqual(['Name', 'Charlie', 'Alpha', 'Bravo', 'Delta'])
    editor.destroy()
  })

  it('sorts numerically descending', async () => {
    const { editor } = await openTable(TABLE5)
    expect(
      sortTableRows(options([level(2, 'number', true)]))(editor.state, editor.view.dispatch),
    ).toBe(true)
    expect(rowNames(editor)).toEqual(['Name', 'Delta', 'Bravo', 'Alpha', 'Charlie'])
    editor.destroy()
  })

  it('puts unparseable keys last ascending and first descending (BUG-744)', async () => {
    // BUG-744 decision: `descending` negates the whole comparison, including
    // the unparseable-vs-number part — Word/Excel order text above numbers in
    // a descending numeric column. The comment in sort.ts used to claim "in
    // both directions" and the PAR-102 report claimed the same; the code (and
    // this pin) is the specification.
    const grid = '<w:tblGrid>' + '<w:gridCol w:w="4000"/>'.repeat(1) + '</w:tblGrid>'
    const rows = ['5', 'n/a', '1300'].map((v) => tr([v]))
    const xml = `<w:tbl><w:tblPr/>${grid}${tr(['Num'], true)}${rows.join('')}</w:tbl>`
    const asc = await openTable(xml)
    expect(
      sortTableRows(options([level(0, 'number')]))(asc.editor.state, asc.editor.view.dispatch),
    ).toBe(true)
    expect(rowNames(asc.editor)).toEqual(['Num', '5', '1300', 'n/a'])
    asc.editor.destroy()
    const desc = await openTable(xml)
    expect(
      sortTableRows(options([level(0, 'number', true)]))(
        desc.editor.state,
        desc.editor.view.dispatch,
      ),
    ).toBe(true)
    expect(rowNames(desc.editor)).toEqual(['Num', 'n/a', '1300', '5'])
    desc.editor.destroy()
  })

  it('sorts by date descending', async () => {
    const { editor, parsed } = await openTable(TABLE5)
    expect(
      sortTableRows(options([level(3, 'date', true)]))(editor.state, editor.view.dispatch),
    ).toBe(true)
    expect(rowNames(editor)).toEqual(['Name', 'Delta', 'Charlie', 'Alpha', 'Bravo'])
    expect(
      (await parseDocx(await resaved(editor, parsed))).blocks[0].table!.rows.map((r) =>
        r[0].paras.join(''),
      ),
    ).toEqual(['Name', 'Delta', 'Charlie', 'Alpha', 'Bravo'])
    editor.destroy()
  })

  it('sorts multi-level: city ascending then name ascending', async () => {
    const { editor } = await openTable(TABLE5)
    expect(
      sortTableRows(options([level(1, 'text'), level(0, 'text')]))(
        editor.state,
        editor.view.dispatch,
      ),
    ).toBe(true)
    expect(rowNames(editor)).toEqual(['Name', 'Alpha', 'Charlie', 'Bravo', 'Delta'])
    editor.destroy()
  })

  it('sorts the header row too when header row is off', async () => {
    const { editor } = await openTable(TABLE5)
    expect(
      sortTableRows(options([level(0, 'text')], false))(editor.state, editor.view.dispatch),
    ).toBe(true)
    expect(rowNames(editor)).toEqual(['Alpha', 'Bravo', 'Charlie', 'Delta', 'Name'])
    editor.destroy()
  })

  it('undo restores the original order in a single step', async () => {
    const { editor } = await openTable(TABLE5)
    const before = rowNames(editor)
    expect(sortTableRows(options([level(0, 'text')]))(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    expect(editor.can().undo()).toBe(true)
    editor.commands.undo()
    expect(rowNames(editor)).toEqual(before)
    expect(editor.can().undo()).toBe(false)
    editor.destroy()
  })

  it('refuses to sort a vertically merged table', async () => {
    const merged =
      '<w:tbl><w:tblPr/>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr>' +
      '<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const { editor } = await openTable(merged)
    const scope = sortScope(editor.state)
    expect(scope).toMatchObject({ kind: 'table', hasVerticalMerge: true })
    expect(sortTableRows(options([level(0, 'text')]))(editor.state, editor.view.dispatch)).toBe(
      false,
    )
    expect(rowNames(editor)).toEqual(['B', 'c', 'A'])
    editor.destroy()
  })

  it('declines a no-op sort without touching history', async () => {
    const { editor } = await openTable(TABLE5)
    expect(sortTableRows(options([]))(editor.state, editor.view.dispatch)).toBe(false)
    expect(sortTableRows(options([level(0, 'text')]))(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    // already sorted: the repeat sort is a no-op and must not add an undo step
    expect(sortTableRows(options([level(0, 'text')]))(editor.state, editor.view.dispatch)).toBe(
      false,
    )
    expect(editor.can().undo()).toBe(true)
    editor.commands.undo()
    expect(rowNames(editor)).toEqual(['Name', 'Delta', 'Alpha', 'Charlie', 'Bravo'])
    expect(editor.can().undo()).toBe(false)
    editor.destroy()
  })
})

/* ================= paragraph sorting ================= */

const PARA_DOC =
  '<w:p><w:r><w:t>pear</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Apple</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>cherry</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>banana</w:t></w:r></w:p>'

async function openDoc(bodyXml: string): Promise<{ editor: Editor; parsed: ParsedDocFull }> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

/** select across the top-level blocks (from inside the first to inside the last) */
function selectBlocks(editor: Editor, first: number, last: number) {
  const doc = editor.state.doc
  let from = 1
  let to = 0
  for (let i = 0; i <= last; i++) {
    if (i === first) from = to + 1
    to += doc.child(i).nodeSize
  }
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(doc, from, to - 1)))
}

function paragraphTexts(editor: Editor): string[] {
  const texts: string[] = []
  editor.state.doc.forEach((block) =>
    texts.push(block.type.name === 'docParagraph' ? block.textContent : '<table>'),
  )
  return texts
}

describe('paragraph sorting', () => {
  it('sorts the selected paragraphs by text, case-insensitively', async () => {
    const { editor, parsed } = await openDoc(PARA_DOC)
    selectBlocks(editor, 0, 3)
    expect(sortScope(editor.state)).toMatchObject({ kind: 'paragraphs', count: 4 })
    expect(sortSelectedParagraphs([level(0, 'text')])(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    expect(paragraphTexts(editor)).toEqual(['Apple', 'banana', 'cherry', 'pear'])
    // byte-identical paragraphs in the new order: the saved body follows it
    const reparsed = await parseDocx(await resaved(editor, parsed))
    const texts = reparsed.blocks
      .filter((b) => b.type === 'paragraph')
      .map((b) => b.runs?.map((r) => r.text).join('') ?? '')
    expect(texts).toEqual(['Apple', 'banana', 'cherry', 'pear'])
    editor.destroy()
  })

  it('sorts descending and undoes in one step', async () => {
    const { editor } = await openDoc(PARA_DOC)
    selectBlocks(editor, 0, 3)
    expect(
      sortSelectedParagraphs([level(0, 'text', true)])(editor.state, editor.view.dispatch),
    ).toBe(true)
    expect(paragraphTexts(editor)).toEqual(['pear', 'cherry', 'banana', 'Apple'])
    expect(editor.can().undo()).toBe(true)
    editor.commands.undo()
    expect(paragraphTexts(editor)).toEqual(['pear', 'Apple', 'cherry', 'banana'])
    expect(editor.can().undo()).toBe(false)
    editor.destroy()
  })

  it('sorts only the paragraphs the selection touches', async () => {
    const { editor } = await openDoc(PARA_DOC)
    selectBlocks(editor, 0, 1)
    expect(sortScope(editor.state)).toMatchObject({ kind: 'paragraphs', count: 2 })
    expect(sortSelectedParagraphs([level(0, 'text')])(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    expect(paragraphTexts(editor)).toEqual(['Apple', 'pear', 'cherry', 'banana'])
    editor.destroy()
  })

  it('refuses a selection containing a table', async () => {
    const { editor } = await openDoc(
      `<w:p><w:r><w:t>b</w:t></w:r></w:p>${TABLE5}<w:p><w:r><w:t>a</w:t></w:r></w:p>`,
    )
    // caret inside the table: the table path owns it
    expect(sortScope(editor.state)).toBeNull()
    selectBlocks(editor, 0, 2)
    expect(sortSelectedParagraphs([level(0, 'text')])(editor.state, editor.view.dispatch)).toBe(
      false,
    )
    expect(paragraphTexts(editor)).toEqual(['b', '<table>', 'a'])
    editor.destroy()
  })

  it('refuses a collapsed selection or a single paragraph', async () => {
    const { editor } = await openDoc(PARA_DOC)
    expect(sortSelectedParagraphs([level(0, 'text')])(editor.state, editor.view.dispatch)).toBe(
      false,
    )
    selectBlocks(editor, 2, 2)
    expect(sortScope(editor.state)).toBeNull()
    editor.destroy()
  })
})

/* ================= Sort dialog ================= */

import { createElement, act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { SortDialog } from '../src/renderer/components/SortDialog'

function renderDialog(editor: Editor): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(SortDialog, { editor, onClose: () => {} })))
  return { container, root }
}

// the dialog renders through the React locale context, which defaults to zh
describe('Sort dialog', () => {
  it('shows the table scope and applies the sort on OK', async () => {
    const { editor } = await openTable(TABLE5)
    const { container, root } = renderDialog(editor)
    expect(container.querySelector('.sort-scope')!.textContent).toBe('表格:5 行 × 5 列')
    act(() => {
      ;(container.querySelector('.modal-actions .btn-primary') as HTMLButtonElement).click()
    })
    expect(rowNames(editor)).toEqual(['Name', 'Alpha', 'Bravo', 'Charlie', 'Delta'])
    act(() => root.unmount())
    editor.destroy()
  })

  it('refuses a merged table: warning shown, OK disabled', async () => {
    const merged =
      '<w:tbl><w:tblPr/>' +
      '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr>' +
      '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const { editor } = await openTable(merged)
    const { container, root } = renderDialog(editor)
    expect(container.querySelector('.sort-warn')!.textContent).toBe(
      '无法排序:暂不支持含合并单元格的表格',
    )
    expect(
      (container.querySelector('.modal-actions .btn-primary') as HTMLButtonElement).disabled,
    ).toBe(true)
    act(() => root.unmount())
    editor.destroy()
  })

  it('sorts the selected paragraphs from the dialog', async () => {
    const { editor } = await openDoc(PARA_DOC)
    selectBlocks(editor, 0, 3)
    const { container, root } = renderDialog(editor)
    expect(container.querySelector('.sort-scope')!.textContent).toBe('选中的段落:4 段')
    act(() => {
      ;(container.querySelector('.modal-actions .btn-primary') as HTMLButtonElement).click()
    })
    expect(paragraphTexts(editor)).toEqual(['Apple', 'banana', 'cherry', 'pear'])
    act(() => root.unmount())
    editor.destroy()
  })
})

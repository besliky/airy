/**
 * Dirty-HTML robustness (BUG-1650): inline formatting tags left open across
 * implicit block boundaries (<h1>… <b>…</h1><p>…, unclosed <p>/<li> after an
 * open <b>/<i>/<span>) used to collapse everything after the first heading
 * into one run-paragraph — lists and tables vanished as structures. The
 * in-page normalization splits those misnested inline wrappers before
 * classification, so paragraph/list/table structures survive; well-formed
 * documents are untouched. One Chrome is shared by every case in this file
 * (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

interface ShapeNode {
  type: string
  text?: string
  items?: number
  rows?: number
}

/** Type/text outline of the extracted IR (docsettings dropped). */
async function irShape(html: string, name: string): Promise<ShapeNode[]> {
  const { ir } = await convertHtml(html, name)
  return ir
    .filter((node: any) => node.type !== 'docsettings')
    .map((node: any) => ({
      type: node.type,
      text: (node.runs || []).map((run: any) => run.text).join(''),
      items: node.items?.length,
      rows: node.rows?.length,
    }))
}

test('unclosed bold inside h1 keeps paragraphs, list, and table (audit vector)', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <h1>Dirty <b>unclosed bold</h1>
      <p>Paragraph one
      <p>Paragraph two never closed
      <ul><li>alpha<li>beta</ul>
      <table><tr><td>no thead</td><td>cells</td></tr></table>
    </body></html>`,
    'unclosed-bold-heading',
  )
  assert.equal(shape[0].type, 'heading')
  const paras = shape.filter((node) => node.type === 'para')
  assert.ok(
    paras.some((node) => node.text?.includes('Paragraph one')),
    'first paragraph should survive as its own block',
  )
  assert.ok(
    paras.some((node) => node.text?.includes('Paragraph two never closed')),
    'second paragraph should survive as its own block',
  )
  const list = shape.find((node) => node.type === 'list')
  assert.ok(list, 'list should survive as a structure')
  assert.equal(list!.items, 2)
  const table = shape.find((node) => node.type === 'table')
  assert.ok(table, 'table should survive as a structure')
  assert.ok(shape.length <= 6, 'content must not duplicate across blocks')
})

test('open b before unclosed p keeps paragraphs and native table', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <b>lead-in
      <p>para one
      <p>para two
      <table><tr><td>cell</td></tr></table>
    </body></html>`,
    'open-b-unclosed-p',
  )
  const paras = shape.filter((node) => node.type === 'para')
  assert.ok(paras.length >= 3, `lead-in and both paragraphs expected, got ${JSON.stringify(shape)}`)
  assert.ok(
    shape.some((node) => node.type === 'table'),
    'table should survive as a structure',
  )
})

test('unclosed li with open em keeps the list, following paragraph, and table', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <ul><li>item <em>accent</ul>
      <p>after list
      <table><tr><td>c1</td><td>c2</td></tr></table>
    </body></html>`,
    'unclosed-li-open-em',
  )
  const list = shape.find((node) => node.type === 'list')
  assert.ok(list, 'list should survive as a structure')
  assert.ok(shape.some((node) => node.type === 'para' && node.text?.includes('after list')))
  assert.ok(shape.some((node) => node.type === 'table'))
})

test('nested unclosed b and i keep paragraphs and list', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <p>text <b>bold <i>both
      <p>after open tags
      <ul><li>n1<li>n2</ul>
    </body></html>`,
    'nested-unclosed',
  )
  const paras = shape.filter((node) => node.type === 'para')
  assert.ok(
    paras.some((node) => node.text?.includes('text')),
    'first paragraph should survive',
  )
  assert.ok(
    paras.some((node) => node.text?.includes('after open tags')),
    'second paragraph should survive',
  )
  const list = shape.find((node) => node.type === 'list')
  assert.ok(list && list.items === 2, 'list should survive as a structure')
})

test('unclosed strong in h1 keeps structure across the following h2 and table', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <h1>H <strong>strong</h1>
      <p>para
      <h2>Sub</h2>
      <table><tr><td>t1</td><td>t2</td></tr></table>
    </body></html>`,
    'unclosed-before-h2',
  )
  assert.equal(shape.filter((node) => node.type === 'heading').length, 2)
  assert.ok(shape.some((node) => node.type === 'para' && node.text?.includes('para')))
  assert.ok(
    shape.some((node) => node.type === 'table'),
    'table should survive as a structure',
  )
})

test('mixed dirty html exports native list and table structures', async () => {
  const { xml } = await convertHtml(
    `<!doctype html>
<html>
<body>
  <h1 style='color: rgb(10, 20, 30)'>Dirty <b>unclosed bold</h1>
  <p class="outer" data-x='say "hi"'>Paragraph one
  <p>Paragraph two never closed
  <ul>
    <li>alpha
    <li>beta <span style="color:red">red tail
  </ul>
  <table><tr><td>no thead</td><td>cells</td></tr></table>
  <custom-element>shadow?</custom-element>
  <div style="background: url('a.jpg')">bg</div>
</body>`,
    'mixed-dirty',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml!, /<w:tbl>/, 'table must export as a native w:tbl')
  assert.match(documentXml!, /w:numPr/, 'list must export with numbering')
  const paragraphCount = (documentXml!.match(/<w:p[ >]/g) || []).length
  assert.ok(paragraphCount >= 4, `expected separate paragraphs, got ${paragraphCount}`)
})

test('well-formed heading, paragraphs, list, and table are unchanged', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <h1>Clean <b>title</b></h1>
      <p>First</p>
      <p>Second</p>
      <ul><li>one</li><li>two</li></ul>
      <table><tr><td>a</td><td>b</td></tr></table>
    </body></html>`,
    'valid-basic',
  )
  assert.deepEqual(
    shape.map((node) => node.type),
    ['heading', 'para', 'para', 'list', 'table'],
  )
  assert.equal(shape[1].text, 'First')
  assert.equal(shape[2].text, 'Second')
  assert.equal(shape[3].items, 2)
  assert.equal(shape[4].rows, 1)
})

test('well-formed bold stays bound to its own words', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><body>
      <h1>Clean <b>title</b></h1>
      <p>plain <b>bold</b> tail</p>
    </body></html>`,
    'valid-bold',
  )
  const heading = ir.find((node: any) => node.type === 'heading')
  assert.deepEqual(
    heading.runs.map((run: any) => run.text),
    ['Clean ', 'title'],
  )
  const para = ir.find((node: any) => node.type === 'para')
  assert.deepEqual(
    para.runs.map((run: any) => [run.text, Boolean(run.bold)]),
    [
      ['plain ', false],
      ['bold', true],
      [' tail', false],
    ],
  )
})

test('valid nested lists keep their levels', async () => {
  const shape = await irShape(
    `<!doctype html><html><body>
      <ul><li>outer<ul><li>inner</li></ul></li><li>second</li></ul>
    </body></html>`,
    'valid-nested-list',
  )
  const lists = shape.filter((node) => node.type === 'list')
  assert.ok(lists.length >= 1, 'outer list should survive')
  const totalItems = lists.reduce((sum, node) => sum + (node.items || 0), 0)
  assert.equal(totalItems, 3, 'inner list items should survive')
})

test('valid table with header row exports all cells', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <table>
        <thead><tr><th>H1</th><th>H2</th></tr></thead>
        <tbody><tr><td>v1</td><td>v2</td></tr></tbody>
      </table>
    </body></html>`,
    'valid-table',
  )
  const documentXml = await xml('word/document.xml')
  assert.equal((documentXml!.match(/<w:tc>/g) || []).length, 4)
  assert.match(documentXml!, /H1/)
  assert.match(documentXml!, /v2/)
})

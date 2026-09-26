// Regression tests for BUG-1777: read_document renders nested tables and
// field-cache values, so the agent sees exactly the content findReplace can
// reach (edit-side visibility used to exceed read-side visibility).
import { execFileSync } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import type { Block } from '@airy-office/docx-engine'

import { DocxSession } from '../src/docx/session.js'
import { blockPreviewText, blockToHtml } from '../src/docx/html.js'
import { buildBodyDocx } from './helpers/docx-fixture.js'

let root: string

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-nested-'))
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

// outer 2x2: texts, a nested inner 2x2 (with an in-paragraph fldSimple) and an
// in-paragraph complex field; those are the field forms the engine models
// (direct w:tc-level field bytes are engine parse scope, invisible to the
// model on BOTH the read and the edit side)
const NESTED_BODY_XML = [
  '<w:p><w:r><w:t>Intro paragraph.</w:t></w:r></w:p>',
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>OUTER head</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>OUTER tail</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p></w:p>' +
    // nested 2x2 table inside the outer cell
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>INNER-NEEDLE head</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>inner b</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>inner c</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:fldSimple w:instr="=SUM(LEFT)"><w:r><w:t>30</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr>' +
    '</w:tbl>' +
    '</w:tc>' +
    // complex field formula inside the outer cell's paragraph
    '<w:tc><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> =SUM(LEFT) </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>33</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl>',
].join('')

async function openNestedSession(): Promise<DocxSession> {
  const docPath = join(root, 'nested.docx')
  await writeFile(docPath, await buildBodyDocx(NESTED_BODY_XML))
  return DocxSession.open(docPath, root)
}

describe('read_document renders nested tables and field caches (BUG-1777)', () => {
  it('shows nested-table text and formula values instead of empty cells', async () => {
    const session = await openNestedSession()
    const html = session.readDocument({ blocks: [1] })
    // the nested table's text is visible...
    expect(html).toContain('INNER-NEEDLE head')
    expect(html).toContain('inner c')
    // ...as a real <table> nested inside the outer cell
    expect(html).toContain('<td><table><tr><td>INNER-NEEDLE head</td>')
    // formula cells show the cached value with the instruction note
    expect(html).toContain('<td>30 [formula: =SUM(LEFT)]</td>')
    expect(html).toContain('<td>33 [formula: =SUM(LEFT)]</td>')
    // the outer table's structure is intact: two <table> elements (outer +
    // inner), four rows total, the answer opens with the outer table
    expect(html).toContain('Selected block content (restricted HTML):\n<table><tr>')
    expect(html.trimEnd().endsWith('</table>')).toBe(true)
    expect(html.match(/<table>/g)).toHaveLength(2)
    expect(html.match(/<tr>/g)).toHaveLength(4)
    await session.close()
  })

  it('shows nested text and formula values in the overview preview too', async () => {
    const session = await openNestedSession()
    const overview = session.readDocument()
    const line = overview.split('\n').find((l) => l.startsWith('1|table|'))
    // nested text rides the (60-char-clipped) preview line; no raw
    // instruction-plus-cache noise for the complex-field cell
    expect(line).toContain('INNER-NEEDLE head')
    expect(line).not.toContain('=SUM(LEFT) 33')
    await session.close()
  })

  it('read shows exactly what findReplace replaces (audit scenario)', async () => {
    const session = await openNestedSession()
    // the nested needle the read just displayed is replaceable
    const { results } = session.applyOps([
      { op: 'findReplace', find: 'INNER-NEEDLE', replace: 'INNER-HIT' },
    ])
    expect(results[0]).toMatchObject({ changed: 1 })
    // read-after-write: the agent sees its own edit inside the nested table
    expect(session.readDocument({ blocks: [1] })).toContain('INNER-HIT')

    // the complex-field cache the read displayed is editable text too
    const cache = session.applyOps([{ op: 'findReplace', find: '33', replace: '34' }])
    expect(cache.results[0]).toMatchObject({ changed: 1 })
    expect(session.readDocument({ blocks: [1] })).toContain('<td>34 [formula: =SUM(LEFT)]</td>')

    const saved = await session.save()
    expect(saved.unchanged).toBe(false)
    // both tables survive the save; the nested fldSimple keeps its original
    // form and the edited complex field re-emits as a second w:fldSimple
    const xml = execFileSync('unzip', ['-p', saved.path, 'word/document.xml'], {
      encoding: 'utf8',
    })
    expect(xml.match(/<w:tbl>/g)).toHaveLength(2)
    expect(xml.match(/<w:fldSimple/g)).toHaveLength(2)
    expect(xml).toContain('w:instr="=SUM(LEFT)"')
    expect(xml).toContain('INNER-HIT')
    // reopen: the nested model carries the replacement
    const { parseDocx } = await import('@airy-office/docx-engine')
    const reparsed = await parseDocx(new Uint8Array(await readFile(saved.path)))
    const outer = reparsed.blocks.find((b) => b.type === 'table')?.table
    expect(outer?.rows[1]?.[0]?.nestedTables?.[0]?.rows[0]?.[0]?.paras).toEqual(['INNER-HIT head'])
    await session.close()
  })

  it('renders hand-built nested models with annotated field runs', () => {
    // whitespace-only formula caches (the engine's empty-cache placeholder)
    // still announce themselves instead of reading as an empty cell
    const model = {
      rows: [
        [
          {
            paras: ['host'],
            nestedTables: [
              {
                rows: [
                  [
                    {
                      paras: ['deep'],
                      richParas: [{ runs: [{ text: ' ', formulaField: '=SUM(ABOVE)' }] }],
                    },
                  ],
                ],
              },
            ],
            nestedTableAnchors: [1],
          },
        ],
      ],
    }
    const block = { type: 'table', table: model } as unknown as Block
    expect(blockToHtml(block)).toBe(
      '<table><tr><td>host\n<table><tr><td>[formula: =SUM(ABOVE)]</td></tr></table></td></tr></table>',
    )
    expect(blockPreviewText(block)).toBe('host [formula: =SUM(ABOVE)]')
  })
})

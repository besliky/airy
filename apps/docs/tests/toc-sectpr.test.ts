/**
 * BUG-1003: updating a TOC/ToF field must carry over the section break that
 * rides in the pPr of the region's last paragraph. Regenerating the field
 * entries used to drop the sectPr, silently destroying the layout of the
 * section that follows the TOC (columns, margins, footer references).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { updateTocField } from '../src/renderer/components/ribbon-references-tab'

/** TOC field whose last entry paragraph closes the field AND the section */
const TOC_WITH_SECTPR_XML =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:hyperlink w:anchor="_Toc000001" w:history="1"><w:r><w:t>Stale entry</w:t></w:r></w:hyperlink></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:noProof/></w:rPr>' +
  '<w:sectPr><w:footerReference r:id="rId9" w:type="default"/>' +
  '<w:pgSz w:w="15840" w:h="12240"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr>' +
  '<w:hyperlink w:anchor="_Toc000002" w:history="1"><w:r><w:t>Stale entry two</w:t></w:r></w:hyperlink>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

const HEADINGS_XML =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter One</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Deeper</w:t></w:r></w:p>'

async function openTocDoc() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TOC_WITH_SECTPR_XML + HEADINGS_XML }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

/** genXml of every regenerated toc line in document order */
function tocLineXmls(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) => {
    const field = node.attrs.fieldDisplay as { kind?: string } | null
    if (field?.kind === 'tocLine') out.push(String(node.attrs.genXml ?? ''))
  })
  return out
}

/** body XML in editor order from a save plan */
function planBodyXml(
  plan: ReturnType<typeof pmDocToSavePlan>,
  originals: Array<{ docxIndex: number | null; originalXml?: string | null }>,
): string {
  let xml = ''
  for (const b of plan.saveBlocks) {
    if (b.kind === 'xml') xml += b.xml
    else if (b.kind === 'original')
      xml += originals.find((x) => x.docxIndex === b.docxIndex)?.originalXml ?? ''
  }
  return xml
}

describe('updateTocField keeps the region trailing section break (BUG-1003)', () => {
  it('re-attaches the last paragraph sectPr to the regenerated last entry', async () => {
    const { editor, parsed } = await openTocDoc()
    expect(updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })).toBe(
      'updated',
    )
    const lines = tocLineXmls(editor)
    expect(lines).toHaveLength(2) // stale entries replaced by the real headings
    expect(lines[0]).toContain('Chapter One')
    expect(lines[0]).not.toContain('<w:sectPr')
    // the section break survives on the LAST regenerated entry, inside its pPr
    expect(lines[1]).toContain('Deeper')
    expect(lines[1]).toContain('<w:pgSz w:w="15840" w:h="12240"/>')
    expect(lines[1]).toMatch(/<w:sectPr>[\s\S]*<\/w:sectPr><\/w:pPr>/)
  })

  it('the preserved sectPr reaches the saved body (no section loss on save)', async () => {
    const { editor, parsed } = await openTocDoc()
    updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })
    const body = planBodyXml(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks),
      parsed.blocks,
    )
    expect(body).toContain('<w:footerReference r:id="rId9" w:type="default"/>')
    // exactly one paragraph-level sectPr in the plan body: the trailing
    // body-level sectPr is a hidden block the save pipeline appends on its own
    expect(body.match(/<w:sectPr[ >]/g)).toHaveLength(1)
  })
})

/** a two-column TOC: a continuous mid-region sectPr closes the first column's
 *  section, the trailing one closes the region — both must survive an update */
const TOC_TWO_COLUMNS_XML =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:sectPr><w:type w:val="continuous"/>' +
  '<w:cols w:num="2" w:space="360"/></w:sectPr></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:hyperlink w:anchor="_Toc000001" w:history="1"><w:r><w:t>Stale col one</w:t></w:r></w:hyperlink></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  '<w:hyperlink w:anchor="_Toc000002" w:history="1"><w:r><w:t>Stale col two</w:t></w:r></w:hyperlink></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:sectPr>' +
  '<w:footerReference r:id="rId9" w:type="default"/>' +
  '<w:pgSz w:w="15840" w:h="12240"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr>' +
  '<w:hyperlink w:anchor="_Toc000003" w:history="1"><w:r><w:t>Stale tail</w:t></w:r></w:hyperlink>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

async function openTwoColumnTocDoc() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TOC_TWO_COLUMNS_XML + HEADINGS_XML }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('updateTocField keeps mid-region continuous section breaks (BUG-1107)', () => {
  it('re-attaches every sectPr of the region, in document order', async () => {
    const { editor, parsed } = await openTwoColumnTocDoc()
    expect(updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })).toBe(
      'updated',
    )
    const lines = tocLineXmls(editor)
    expect(lines).toHaveLength(2)
    // the continuous column break lands on an intermediate regenerated entry…
    expect(lines[0]).toContain('Chapter One')
    expect(lines[0]).toContain('<w:type w:val="continuous"/>')
    expect(lines[0]).toContain('<w:cols w:num="2" w:space="360"/>')
    // …the trailing section break stays on the last one (BUG-1003 behavior)
    expect(lines[1]).toContain('Deeper')
    expect(lines[1]).toContain('<w:footerReference r:id="rId9" w:type="default"/>')
    expect(lines[1]).not.toContain('w:continuous')
  })

  it('both sections reach the saved body', async () => {
    const { editor, parsed } = await openTwoColumnTocDoc()
    updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })
    const body = planBodyXml(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks),
      parsed.blocks,
    )
    expect(body).toContain('<w:cols w:num="2" w:space="360"/>')
    expect(body).toContain('<w:footerReference r:id="rId9" w:type="default"/>')
    // exactly the two region sectPrs (the trailing body-level sectPr is a
    // hidden block the save pipeline appends on its own)
    expect(body.match(/<w:sectPr[ >]/g)).toHaveLength(2)
  })

  it('more breaks than regenerated entries keep their own break paragraphs', async () => {
    // three sectPrs over a region that regenerates to a single entry line:
    // one attaches, the surplus stays as standalone section-break paragraphs
    const xml =
      '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:cols w:num="2"/></w:sectPr></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h \\z \\u </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:hyperlink w:anchor="_Toc000001" w:history="1"><w:r><w:t>Stale</w:t></w:r></w:hyperlink></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:cols w:num="3"/></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240"/></w:sectPr></w:pPr>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          xml + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Solo</w:t></w:r></w:p>',
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    expect(updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })).toBe(
      'updated',
    )
    const body = planBodyXml(
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks),
      parsed.blocks,
    )
    expect(body).toContain('<w:cols w:num="2"/>')
    expect(body).toContain('<w:cols w:num="3"/>')
    expect(body).toContain('<w:pgSz w:w="15840" w:h="12240"/>')
    expect(body.match(/<w:sectPr[ >]/g)).toHaveLength(3)
  })
})

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
    expect(
      updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true }),
    ).toBe('updated')
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
    const body = planBodyXml(pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks), parsed.blocks)
    expect(body).toContain('<w:footerReference r:id="rId9" w:type="default"/>')
    // exactly one paragraph-level sectPr in the plan body: the trailing
    // body-level sectPr is a hidden block the save pipeline appends on its own
    expect(body.match(/<w:sectPr[ >]/g)).toHaveLength(1)
  })
})

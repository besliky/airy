import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@airy-office/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const BODY =
  '<w:p><w:bookmarkStart w:id="1" w:name="Conclusion"/><w:bookmarkEnd w:id="1"/>' +
  '<w:r><w:t>Conclusion paragraph.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Plain paragraph.</w:t></w:r></w:p>'

async function open() {
  const source = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

async function docXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

describe('bookmarks & cross-references in the editor', () => {
  it('renders data-bookmarks and keeps an untouched doc byte-identical', async () => {
    const { editor, parsed, source } = await open()
    expect(editor.view.dom.querySelector('[data-bookmarks~="Conclusion"]')?.textContent).toBe(
      'Conclusion paragraph.',
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('adding a bookmark regenerates the paragraph with bookmarkStart', async () => {
    const { editor, parsed } = await open()
    // second paragraph node position
    let secondPos = -1
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'docParagraph' && node.textContent === 'Plain paragraph.')
        secondPos = offset
    })
    const node = editor.state.doc.nodeAt(secondPos)!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(secondPos, undefined, {
        ...node.attrs,
        bookmarks: ['NewBookmark'],
      }),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[1].bookmarks).toEqual(['NewBookmark'])
    expect(reparsed.blocks[0].bookmarks).toEqual(['Conclusion'])
    editor.destroy()
  })

  it('inserting a cross-reference saves a REF field pointing at the bookmark', async () => {
    const { editor, parsed, source } = await open()
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.near(editor.state.doc.resolve(editor.state.doc.content.size - 1)),
      ),
    )
    editor
      .chain()
      .insertContent({
        type: 'text',
        text: 'Conclusion paragraph.',
        marks: [{ type: 'refField', attrs: { name: 'Conclusion' } }],
      })
      .run()
    expect(editor.view.dom.querySelector('span[data-ref-field="Conclusion"]')).toBeTruthy()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(await docXmlOf(saved)).toContain('REF Conclusion \\h')
    const reparsed = await parseDocx(saved)
    const refRun = reparsed.blocks[1].runs!.find((r) => r.refField === 'Conclusion')
    expect(refRun?.text).toBe('Conclusion paragraph.')
    editor.destroy()
  })

  it('editing a paragraph with a switched REF (\\p) keeps the instruction verbatim', async () => {
    const body =
      BODY +
      '<w:p><w:r><w:t>See p.</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> REF Conclusion \\p \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>1</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const source = await buildDocx({ bodyXml: body })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    // the mark carries the original instruction through the editor model
    const marks: Array<{ type: { name: string }; attrs: Record<string, unknown> }> = []
    editor.state.doc.descendants((node) => {
      for (const m of node.marks) marks.push(m as never)
    })
    const mark = marks.find((m) => m.type.name === 'refField')
    expect(mark?.attrs.instr).toBe(' REF Conclusion \\p \\h ')
    // edit the paragraph (append text after the field) and save
    const end = editor.state.doc.content.size
    editor.chain().insertContentAt(end, { type: 'text', text: '!' }).run()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBeGreaterThan(0)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await docXmlOf(saved)
    expect(xml).toContain(' REF Conclusion \\p \\h ')
    expect(xml).toContain('<w:t>1</w:t>')
    editor.destroy()
  })
})

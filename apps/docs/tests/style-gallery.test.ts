/**
 * PAR-114 style gallery: applying every gallery paragraph style — Normal, the
 * nine built-in heading levels and document custom styles — must land as the
 * right w:pStyle in the saved XML and survive a reopen (round-trip).
 */
import { afterEach, describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  pmNodeToGeneratedBlock,
  type PmNode,
} from '../src/renderer/editor/convert'
import {
  applyDocumentParagraphStyle,
  applyParagraphStyle,
  builtinHeadingStyleId,
  retagStyleParagraphs,
} from '../src/renderer/components/ribbon-tabs'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const CUSTOM_STYLES =
  '<w:style w:type="paragraph" w:styleId="MyQuote"><w:name w:val="My Quote"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ChapterTitle"><w:name w:val="Chapter Title"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>'

const BODY = ['one', 'two', 'three'].map((x) => `<w:p><w:r><w:t>${x}</w:t></w:r></w:p>`).join('')

afterEach(() => drainTrackedEditors())

async function open(bodyXml = BODY, extraStylesXml = CUSTOM_STYLES) {
  const source = await buildDocx({ bodyXml, extraStylesXml })
  const parsed = await parseDocx(source)
  const editor = createTrackedEditor({
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, source }
}

/** put the caret in the first paragraph, apply, return its generated block */
function applyToFirst(editor: Editor, fn: (editor: Editor) => void) {
  editor.commands.setTextSelection(1)
  fn(editor)
  return pmNodeToGeneratedBlock(editor.state.doc.child(0).toJSON())
}

async function savedFirstParagraphXml(
  editor: Editor,
  parsed: Awaited<ReturnType<typeof open>>['parsed'],
) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const reparsed = await parseDocx(bytes)
  return { xml: reparsed.blocks[0].originalXml ?? '', reparsed, bytes }
}

describe('gallery heading styles h1-h9', () => {
  it.each([1, 2, 3, 4, 5, 6, 7, 8, 9])(
    'h%d applies a docHeading that saves as its pStyle and reopens as the same level',
    async (level) => {
      const { editor, parsed } = await open()
      const block = applyToFirst(editor, (e) => applyParagraphStyle(e, `h${level}` as `h${1}`))
      expect(block.type).toBe('heading')
      expect(block.level).toBe(level)
      expect(block.styleId).toBe(builtinHeadingStyleId(level))
      const { xml, reparsed } = await savedFirstParagraphXml(editor, parsed)
      expect(xml).toContain(`<w:pStyle w:val="Heading${level}"/>`)
      expect(reparsed.blocks[0]).toMatchObject({
        type: 'heading',
        level,
        styleId: `Heading${level}`,
      })
    },
  )

  it('h4 carries the document-resolved styleId when styles.xml defines one', async () => {
    const { editor } = await open()
    const block = applyToFirst(editor, (e) => applyParagraphStyle(e, 'h4', 'Heading4'))
    expect(block).toMatchObject({ type: 'heading', level: 4, styleId: 'Heading4' })
  })

  it('Normal clears the pStyle of a styled paragraph', async () => {
    const { editor, parsed } = await open(
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    )
    const block = applyToFirst(editor, (e) => applyParagraphStyle(e, 'p'))
    expect(block.type).toBe('paragraph')
    expect(block.styleId).toBeUndefined()
    const { xml } = await savedFirstParagraphXml(editor, parsed)
    expect(xml).not.toContain('<w:pStyle')
  })
})

describe('document custom paragraph styles', () => {
  it('applies a custom style to the paragraph and saves it as pStyle', async () => {
    const { editor, parsed } = await open()
    const block = applyToFirst(editor, (e) =>
      applyDocumentParagraphStyle(e, { styleId: 'MyQuote' }),
    )
    expect(block).toMatchObject({ type: 'paragraph', styleId: 'MyQuote' })
    const { xml, reparsed } = await savedFirstParagraphXml(editor, parsed)
    expect(xml).toContain('<w:pStyle w:val="MyQuote"/>')
    expect(reparsed.blocks[0]).toMatchObject({ type: 'paragraph', styleId: 'MyQuote' })
  })

  it('a style with an outline level turns the paragraph into that heading', async () => {
    const { editor, parsed } = await open()
    const block = applyToFirst(editor, (e) =>
      applyDocumentParagraphStyle(e, { styleId: 'ChapterTitle', headingLevel: 1 }),
    )
    expect(block).toMatchObject({ type: 'heading', level: 1, styleId: 'ChapterTitle' })
    const { xml, reparsed } = await savedFirstParagraphXml(editor, parsed)
    expect(xml).toContain('<w:pStyle w:val="ChapterTitle"/>')
    expect(reparsed.blocks[0]).toMatchObject({ type: 'heading', level: 1, styleId: 'ChapterTitle' })
  })

  it('a custom style replaces a heading (back to body paragraph)', async () => {
    const { editor } = await open(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    )
    const block = applyToFirst(editor, (e) =>
      applyDocumentParagraphStyle(e, { styleId: 'MyQuote' }),
    )
    expect(block).toMatchObject({ type: 'paragraph', styleId: 'MyQuote' })
  })

  it('applying a style over a range retags every selected paragraph', async () => {
    const { editor } = await open()
    const size = editor.state.doc.content.size
    editor.commands.setTextSelection({ from: 1, to: size - 1 })
    applyDocumentParagraphStyle(editor, { styleId: 'MyQuote' })
    const blocks = [0, 1, 2].map((i) => pmNodeToGeneratedBlock(editor.state.doc.child(i).toJSON()))
    expect(blocks.every((b) => b.styleId === 'MyQuote')).toBe(true)
  })
})

describe('retagStyleParagraphs (Modify Style outline change)', () => {
  it('moves headings of the style to the new level and back to body text', async () => {
    const { editor } = await open(
      `<w:p><w:pPr><w:pStyle w:val="ChapterTitle"/></w:pPr><w:r><w:t>chapter</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>body</w:t></w:r></w:p>`,
    )
    // parsed ChapterTitle has outlineLvl 0 → heading level 1
    retagStyleParagraphs(editor, 'ChapterTitle', 3)
    expect(pmNodeToGeneratedBlock(editor.state.doc.child(0).toJSON())).toMatchObject({
      type: 'heading',
      level: 3,
      styleId: 'ChapterTitle',
    })
    retagStyleParagraphs(editor, 'ChapterTitle', null)
    expect(pmNodeToGeneratedBlock(editor.state.doc.child(0).toJSON())).toMatchObject({
      type: 'paragraph',
      styleId: 'ChapterTitle',
    })
  })

  it('a newly outlined style promotes its body paragraphs to headings', async () => {
    const { editor } = await open()
    editor.commands.setTextSelection(1)
    applyDocumentParagraphStyle(editor, { styleId: 'MyQuote' })
    retagStyleParagraphs(editor, 'MyQuote', 2)
    expect(pmNodeToGeneratedBlock(editor.state.doc.child(0).toJSON())).toMatchObject({
      type: 'heading',
      level: 2,
      styleId: 'MyQuote',
    })
    // paragraphs of other styles stay untouched
    expect(pmNodeToGeneratedBlock(editor.state.doc.child(1).toJSON()).type).toBe('paragraph')
  })
})

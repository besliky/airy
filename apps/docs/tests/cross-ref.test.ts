import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import {
  generateCaptionXml,
  generateParagraphXml,
  parseDocx,
  saveDocx,
  type GenerateContext,
} from '@airy-office/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  collectCrossRefSources,
  crossRefCache,
  crossRefInstr,
  ensureCaptionAnchor,
  ensureHeadingTocAnchor,
  findAnchorPos,
  refCacheOf,
  type CrossRefSource,
} from '../src/renderer/components/cross-ref'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const HEADING_XML = generateParagraphXml(
  { type: 'heading', level: 1, runs: [{ text: 'Results' }] },
  GEN_CTX,
)
const BOOKMARKED_XML =
  '<w:p><w:bookmarkStart w:id="1" w:name="Conclusion"/><w:bookmarkEnd w:id="1"/>' +
  '<w:r><w:t>Conclusion paragraph.</w:t></w:r></w:p>'
const CAPTION_XML = generateCaptionXml('Figure', 1, 'System architecture')
const CAPTION2_XML = generateCaptionXml('Figure', 2, 'Inputs')
const PLAIN_XML = '<w:p><w:r><w:t>Plain paragraph.</w:t></w:r></w:p>'

async function open(
  bodyXml = HEADING_XML + BOOKMARKED_XML + CAPTION_XML + CAPTION2_XML + PLAIN_XML,
) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const endOf = (text: string): number => {
    let pos = -1
    editor.state.doc.forEach((node, offset) => {
      if (pos === -1 && node.textContent.includes(text)) pos = offset + node.nodeSize
    })
    return pos
  }
  return { editor, parsed, source, endOf }
}

async function docXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

/** insert a REF field like the modal does and return the saved .docx bytes */
async function insertAndSave(
  editor: Editor,
  parsed: Awaited<ReturnType<typeof open>>['parsed'],
  at: number,
  anchor: string,
  type: 'text' | 'page' | 'number',
  cache: string,
): Promise<Uint8Array> {
  editor
    .chain()
    .insertContentAt(at, {
      type: 'text',
      text: cache,
      marks: [{ type: 'refField', attrs: { name: anchor, instr: crossRefInstr(anchor, type) } }],
    })
    .run()
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  return saveDocx(parsed, plan.saveBlocks)
}

describe('cross-reference sources', () => {
  it('collects headings, SEQ captions (with ordinals), and bookmarks', async () => {
    const { editor, parsed } = await open()
    const sources = collectCrossRefSources(editor, parsed.blocks)
    const headings = sources.filter((s) => s.kind === 'heading')
    const captions = sources.filter((s) => s.kind === 'caption')
    const bookmarks = sources.filter((s) => s.kind === 'bookmark')
    expect(headings.map((s) => s.label)).toEqual(['Results'])
    expect(captions.map((s) => s.seqNumber)).toEqual([1, 2])
    expect(captions.map((s) => s.label)).toEqual([
      'Figure 1 System architecture',
      'Figure 2 Inputs',
    ])
    expect(captions.every((s) => s.seqLabel === 'Figure' && s.anchor === null)).toBe(true)
    expect(bookmarks.map((s) => s.label)).toEqual(['Conclusion'])
    editor.destroy()
  })
})

describe('cross-reference instructions and caches', () => {
  const heading: CrossRefSource = {
    kind: 'heading',
    label: 'Results',
    preview: 'Results',
    level: 1,
    anchor: null,
    pos: 0,
  }
  const caption: CrossRefSource = {
    kind: 'caption',
    label: 'Figure 2 Inputs',
    preview: '',
    level: 1,
    anchor: null,
    pos: 4,
    seqLabel: 'Figure',
    seqNumber: 2,
  }

  it('builds Word-style REF instructions per type', () => {
    expect(crossRefInstr('chap1', 'text')).toBe(' REF chap1 \\h ')
    expect(crossRefInstr('chap1', 'page')).toBe(' REF chap1 \\p \\h ')
    expect(crossRefInstr('chap1', 'number')).toBe(' REF chap1 \\r \\h ')
  })

  it('caches target text, page (from pagination), and SEQ number', () => {
    expect(crossRefCache({ ...heading, anchor: 'chap1' }, 'text')).toBe('Results')
    expect(crossRefCache({ ...heading, anchor: 'chap1' }, 'page', () => 3)).toBe('3')
    // page unknown until pagination: blank placeholder, filled by F9 / Word
    expect(crossRefCache({ ...heading, anchor: 'chap1' }, 'page')).toBe(' ')
    expect(crossRefCache(caption, 'number')).toBe('2')
    expect(crossRefCache(caption, 'text')).toBe('Figure 2 Inputs')
    // heading number is only computable when the heading text carries it
    expect(crossRefCache({ ...heading, label: '3.2 Details' }, 'number')).toBe('3.2')
    expect(crossRefCache(heading, 'number')).toBe(' ')
  })
})

describe('cross-reference insertion round-trip', () => {
  it('heading text reference: stamps a _Toc anchor and saves REF … \\h', async () => {
    const { editor, parsed, endOf } = await open()
    // the modal stamps the heading's hidden anchor before inserting the field
    const stamped = ensureHeadingTocAnchor(editor, 0)
    expect(stamped).toMatch(/^_Toc\d+$/)
    const saved = await insertAndSave(
      editor,
      parsed,
      endOf('Plain paragraph.'),
      stamped!,
      'text',
      'Results',
    )
    const xml = await docXmlOf(saved)
    // the anchor regenerates on the heading paragraph
    expect(xml).toContain(`w:name="${stamped}"`)
    expect(xml).toContain(` REF ${stamped} \\h `)
    // and round-trips back into the run model
    const reparsed = await parseDocx(saved)
    const run = reparsed.blocks.flatMap((b) => b.runs ?? []).find((r) => r.refField === stamped)
    expect(run).toMatchObject({ text: 'Results', refField: stamped })
    editor.destroy()
  })

  it('bookmark page reference: saves REF … \\p \\h with a placeholder cache', async () => {
    const { editor, parsed, endOf } = await open()
    const xml = await docXmlOf(
      await insertAndSave(editor, parsed, endOf('Plain paragraph.'), 'Conclusion', 'page', ' '),
    )
    expect(xml).toContain(' REF Conclusion \\p \\h ')
    editor.destroy()
  })

  it('caption number reference: stamps a _Ref anchor on the caption and saves REF … \\r \\h', async () => {
    const { editor, parsed, endOf } = await open()
    const caption = collectCrossRefSources(editor, parsed.blocks).find(
      (s) => s.kind === 'caption' && s.seqNumber === 2,
    )!
    const anchor = ensureCaptionAnchor(editor, parsed.blocks, caption)
    expect(anchor).toMatch(/^_Ref\d+$/)
    const saved = await insertAndSave(
      editor,
      parsed,
      endOf('Plain paragraph.'),
      anchor!,
      'number',
      '2',
    )
    const xml = await docXmlOf(saved)
    // the caption paragraph carries the hidden bookmark…
    expect(xml).toContain(`w:name="${anchor}"`)
    // …and the reference points at it with \r
    expect(xml).toContain(` REF ${anchor} \\r \\h `)
    // round-trip: the REF keeps its switches and the caption keeps its anchor
    const reparsed = await parseDocx(saved)
    const run = reparsed.blocks.flatMap((b) => b.runs ?? []).find((r) => r.refField === anchor)
    expect(run).toMatchObject({ text: '2', refInstr: ` REF ${anchor} \\r \\h ` })
    expect(reparsed.blocks.some((b) => (b.originalXml ?? '').includes(`w:name="${anchor}"`))).toBe(
      true,
    )
    editor.destroy()
  })
})

describe('F9 REF cache recompute', () => {
  it('recomputes text, page (via pagination callback), and SEQ number', async () => {
    const { editor, parsed } = await open()
    const blocks = parsed.blocks
    // text reference to the bookmarked paragraph
    expect(refCacheOf(editor, blocks, ' REF Conclusion \\h ')).toBe('Conclusion paragraph.')
    // page reference: the page comes from pagination, null when unavailable
    expect(refCacheOf(editor, blocks, ' REF Conclusion \\p \\h ', () => 2)).toBe('2')
    expect(refCacheOf(editor, blocks, ' REF Conclusion \\p \\h ')).toBeNull()
    // number reference to the second Figure caption
    const caption = collectCrossRefSources(editor, blocks).find(
      (s) => s.kind === 'caption' && s.seqNumber === 2,
    )!
    const anchor = ensureCaptionAnchor(editor, blocks, caption)
    expect(refCacheOf(editor, blocks, ` REF ${anchor} \\r \\h `)).toBe('2')
    // gone target: cache stays untouched (null)
    expect(refCacheOf(editor, blocks, ' REF Missing \\h ')).toBeNull()
    editor.destroy()
  })

  it('finds anchors stamped into caption XML by position', async () => {
    const { editor, parsed } = await open()
    const caption = collectCrossRefSources(editor, parsed.blocks).find(
      (s) => s.kind === 'caption' && s.seqNumber === 1,
    )!
    const anchor = ensureCaptionAnchor(editor, parsed.blocks, caption)!
    const pos = findAnchorPos(editor.state.doc, parsed.blocks, anchor)
    expect(pos).toBe(caption.pos)
    editor.destroy()
  })
})

describe('cursor insertion via the editor', () => {
  it('inserts a page reference at the caret with the instr mark', async () => {
    const { editor, parsed, source } = await open(BOOKMARKED_XML + PLAIN_XML)
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.near(editor.state.doc.resolve(editor.state.doc.content.size - 1)),
      ),
    )
    const saved = await insertAndSave(
      editor,
      parsed,
      editor.state.selection.from,
      'Conclusion',
      'page',
      '2',
    )
    const xml = await docXmlOf(saved)
    expect(xml).toContain(' REF Conclusion \\p \\h ')
    // the untouched bookmarked paragraph stays byte-identical
    expect(xml).toContain('Conclusion paragraph.')
    const reparsed = await parseDocx(saved)
    const run = reparsed.blocks
      .flatMap((b) => b.runs ?? [])
      .find((r) => r.refField === 'Conclusion')
    expect(run).toMatchObject({ text: '2', refInstr: ' REF Conclusion \\p \\h ' })
    expect(source).toBeTruthy()
    editor.destroy()
  })
})

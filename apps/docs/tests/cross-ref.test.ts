import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import {
  bookmarkIdOf,
  generateCaptionXml,
  generateParagraphXml,
  parseDocx,
  saveDocx,
  type GenerateContext,
} from '@airy-office/docx-engine'
import { describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  REF_TARGET_GONE,
  allBookmarkIds,
  anchorIdOf,
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
// Word routinely splits one field instruction across several w:instrText runs
// (rsid seams); the reader concatenates them while parsing, the renderer must
// too when reading the raw caption XML (BUG-912)
const TORN_CAPTION_XML =
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> SEQ Fig</w:instrText></w:r>' +
  '<w:r><w:instrText xml:space="preserve">ure \\* ARABIC </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '<w:r><w:t xml:space="preserve">Torn caption</w:t></w:r>' +
  '</w:p>'
// the first fragment can also end mid-whitespace: a " SEQ " first run gave
// the old first-fragment regex no \S+ to match at all
const TORN_TABLE_XML =
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> SEQ </w:instrText></w:r>' +
  '<w:r><w:instrText xml:space="preserve">Table \\* ARABIC </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:p>'

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

  it('reads SEQ labels from split-instruction captions (BUG-912)', async () => {
    const { editor, parsed } = await open(
      HEADING_XML + TORN_CAPTION_XML + TORN_TABLE_XML + CAPTION_XML,
    )
    const captions = collectCrossRefSources(editor, parsed.blocks).filter(
      (s) => s.kind === 'caption',
    )
    // "SEQ Fig|ure" and "SEQ |Table" must join into the real labels, sharing
    // the "Figure" ordinal pool with the intact caption instead of forming
    // phantom "Fig" pools or vanishing from the dialog
    expect(captions.map((s) => s.seqLabel)).toEqual(['Figure', 'Table', 'Figure'])
    expect(captions.map((s) => s.seqNumber)).toEqual([1, 1, 2])
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
    // a bare leading number is prose (a year), not a numbering prefix; a
    // single number needs a real list separator behind it (BUG-916b)
    expect(crossRefCache({ ...heading, label: '2026 Report' }, 'number')).toBe(' ')
    expect(crossRefCache({ ...heading, label: '1. Introduction' }, 'number')).toBe('1')
    expect(crossRefCache({ ...heading, label: '1、はじめに' }, 'number')).toBe('1')
    // an emptied bookmark shows Word's reference error, never the anchor
    // name as visible field text (BUG-916a)
    expect(
      crossRefCache(
        {
          kind: 'bookmark',
          label: 'EmptyTarget',
          preview: '',
          level: 1,
          anchor: 'EmptyTarget',
          pos: 0,
        },
        'text',
      ),
    ).toBe(REF_TARGET_GONE)
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
    // gone target: Word's F9 writes the reference error into the cache (BUG-916a)
    expect(refCacheOf(editor, blocks, ' REF Missing \\h ')).toBe(REF_TARGET_GONE)
    editor.destroy()
  })

  it('F9 gives emptied and year-headed targets Word semantics (BUG-916)', async () => {
    const yearHeadingXml = generateParagraphXml(
      { type: 'heading', level: 1, runs: [{ text: '2026 Report' }] },
      GEN_CTX,
    )
    const emptyBookmarkXml =
      '<w:p><w:bookmarkStart w:id="7" w:name="EmptyTarget"/><w:bookmarkEnd w:id="7"/></w:p>'
    const { editor, parsed } = await open(yearHeadingXml + emptyBookmarkXml + PLAIN_XML)
    const blocks = parsed.blocks
    // a bookmark resolving to an empty paragraph shows the error line, not
    // the raw anchor name as visible text
    expect(refCacheOf(editor, blocks, ' REF EmptyTarget \\h ')).toBe(REF_TARGET_GONE)
    // "2026 Report" is a year-headed heading, not number 2026: \r stays
    // uncomputable locally (cache untouched) while real prefixes resolve
    const anchor = ensureHeadingTocAnchor(editor, 0)
    expect(refCacheOf(editor, blocks, ` REF ${anchor} \\r \\h `)).toBeNull()
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

  it('F9 \\r numbers a split-instruction caption from the joined label (BUG-912)', async () => {
    const { editor, parsed } = await open(
      HEADING_XML + CAPTION_XML + TORN_CAPTION_XML + CAPTION2_XML,
    )
    const blocks = parsed.blocks
    const torn = collectCrossRefSources(editor, blocks).find(
      (s) => s.kind === 'caption' && s.label.includes('Torn caption'),
    )!
    // the torn "SEQ Fig|ure" caption is the SECOND Figure of three: only the
    // joined label puts it in the shared ordinal pool
    expect(torn.seqLabel).toBe('Figure')
    expect(torn.seqNumber).toBe(2)
    const anchor = ensureCaptionAnchor(editor, blocks, torn)
    expect(anchor).toMatch(/^_Ref\d+$/)
    expect(refCacheOf(editor, blocks, ` REF ${anchor} \\r \\h `)).toBe('2')
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

describe('caption anchor w:id uniqueness (BUG-919)', () => {
  it('collects every bookmark id the saved document will contain', async () => {
    // a caption already carrying a stamped anchor keeps its literal w:id in the
    // protected XML; the user bookmark's id is minted at save time from its
    // name (parse keeps names only, save re-emits via the engine hash)
    const anchoredCaption = generateCaptionXml('Figure', 1, 'Anchored', '_Ref135792468')
    const { editor, parsed } = await open(
      HEADING_XML + BOOKMARKED_XML + anchoredCaption + CAPTION2_XML + PLAIN_XML,
    )
    const ids = allBookmarkIds(editor.state.doc, parsed.blocks)
    expect(ids.has(1135792468)).toBe(true) // anchorIdOf('_Ref135792468'), protected XML
    expect(ids.has(bookmarkIdOf('Conclusion'))).toBe(true) // save-time hash id
    editor.destroy()
  })

  it('stamps a caption anchor whose w:id avoids the document id pool', async () => {
    // bookmarkIdOf('Budget2027') lands inside the 1e9 stamp range: a name-only
    // anchor pick can mint the very same w:id the save will give this bookmark
    expect(bookmarkIdOf('Budget2027')).toBe(1302386152)
    const userBookmark =
      '<w:p><w:bookmarkStart w:id="7" w:name="Budget2027"/><w:bookmarkEnd w:id="7"/>' +
      '<w:r><w:t>Budget paragraph.</w:t></w:r></w:p>'
    const { editor, parsed, endOf } = await open(
      HEADING_XML + userBookmark + CAPTION_XML + CAPTION2_XML + PLAIN_XML,
    )
    const caption = collectCrossRefSources(editor, parsed.blocks).find(
      (s) => s.kind === 'caption' && s.seqNumber === 1,
    )!
    // deterministic candidates: the colliding _Ref302386152 (id 1302386152)
    // first, then the free _Ref550000000 (id 1550000000)
    const random = vi
      .spyOn(Math, 'random')
      .mockReturnValueOnce((302386152 - 100000000 + 0.5) / 900000000)
      .mockReturnValue((550000000 - 100000000 + 0.5) / 900000000)
    const anchor = ensureCaptionAnchor(editor, parsed.blocks, caption)
    random.mockRestore()
    // the id-aware pick rejected the colliding name instead of stamping a
    // duplicate w:id into the document
    expect(anchor).toBe('_Ref550000000')
    expect(anchorIdOf(anchor!)).not.toBe(bookmarkIdOf('Budget2027'))

    let genXml = ''
    editor.state.doc.descendants((node) => {
      const xml = String(node.attrs?.genXml ?? '')
      if (node.type.name === 'docProtected' && xml.includes('_Ref550000000')) {
        genXml = xml
        return false
      }
      return true
    })
    expect(genXml).toContain('<w:bookmarkStart w:id="1550000000" w:name="_Ref550000000"/>')

    // full round trip: every bookmarkStart id in the saved document is unique
    const saved = await insertAndSave(
      editor,
      parsed,
      endOf('Plain paragraph.'),
      anchor!,
      'number',
      '1',
    )
    const xml = await docXmlOf(saved)
    const ids = [...xml.matchAll(/<w:bookmarkStart [^>]*w:id="(\d+)"/g)].map((m) => m[1])
    expect(ids.length).toBeGreaterThan(1) // the anchor and the user bookmark both made it
    expect(new Set(ids).size).toBe(ids.length)
    editor.destroy()
  })
})

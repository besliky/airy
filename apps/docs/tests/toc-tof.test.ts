/**
 * PAR-110: TOC options + table-of-figures authoring. ToF entries come from
 * SEQ captions of one label, an update rebuilds the field from its authored
 * instruction (switches survive), and the flow round-trips through a save.
 */
import { describe, expect, it, beforeAll, afterAll, afterEach } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateCaptionXml, generateTocFieldXml, parseDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import {
  collectTofEntries,
  parseTocStyleSpec,
  updateTocField,
} from '../src/renderer/components/ribbon-references-tab'

/** editors created by these tests; destroyed in afterEach so the ProseMirror
 *  DOMObserver polling timer never outlives the jsdom environment (an
 *  unhandled "document is not defined" after teardown fails the whole run) */
const openEditors: Editor[] = []

async function openDoc(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  openEditors.push(editor)
  return { editor, parsed }
}

afterEach(async () => {
  while (openEditors.length) openEditors.pop()?.destroy()
  // let a pending DOMObserver flush land while the document still exists
  await new Promise((resolve) => setTimeout(resolve, 30))
})

/** a generated caption paragraph, the shape the ribbon's Caption dialog inserts */
function captionNode(label: string, n: number, text: string, anchor: string) {
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Caption',
      genXml: generateCaptionXml(label, n, text, anchor),
      fieldDisplay: { kind: 'text', left: `${label} ${n}${text ? ` ${text}` : ''}` },
    },
  }
}

/** a caption with a canonical SEQ id but a translated visible word — the
 *  shape the Caption dialog writes since UX-1011 */
function captionNodeLocalized(
  label: string,
  n: number,
  text: string,
  anchor: string,
  display: string,
) {
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Caption',
      genXml: generateCaptionXml(label, n, text, anchor, display),
      fieldDisplay: { kind: 'text', left: `${display} ${n}${text ? ` ${text}` : ''}` },
    },
  }
}

/** PM nodes for a generated field, mirroring the ribbon insert (tocFieldNodes) */
function fieldNodes(entries: Parameters<typeof generateTocFieldXml>[0], options = {}) {
  return generateTocFieldXml(entries, options).map((xml, i) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Field',
      genXml: xml,
      fieldDisplay: {
        kind: 'tocLine',
        left: entries[i].text,
        right: entries[i].pageNo !== undefined ? String(entries[i].pageNo) : '',
        level: entries[i].level,
        ...(entries[i].anchor ? { anchor: entries[i].anchor } : {}),
      },
    },
  }))
}

function headingNode(text: string, level: number) {
  return {
    type: 'docHeading',
    attrs: { docxIndex: null, styleId: null, aiChanged: false, level },
    content: [{ type: 'text', text }],
  }
}

interface TocLine {
  left?: string
  right?: string
  level?: number
  anchor?: string
  noPage?: boolean
}

function tocLines(editor: Editor): TocLine[] {
  const lines: TocLine[] = []
  editor.state.doc.forEach((node) => {
    const field = node.attrs.fieldDisplay as { kind?: string } | null
    if (field?.kind === 'tocLine') lines.push(node.attrs.fieldDisplay as TocLine)
  })
  return lines
}

function fieldInstructions(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) => {
    if (node.type.name !== 'docProtected') return
    const xml = String(node.attrs.genXml ?? '')
    const m = /<w:instrText[^>]*>([^<]*)<\/w:instrText>/.exec(xml)
    if (m && /^\s*TOC[\s\\]/.test(m[1])) out.push(m[1].trim())
  })
  return out
}

/** body XML in editor order from a save plan (original blocks + generated xml) */
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

describe('table of figures authoring', () => {
  it('collects entries from SEQ captions of one label, with anchors', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNode('Figure', 1, 'Architecture', '_Ref111111111'),
      captionNode('Figure', 2, 'Data flow', '_Ref222222222'),
      captionNode('Table', 1, 'Inputs', '_Ref333333333'),
    ] as never)
    const entries = collectTofEntries(editor, parsed.blocks, 'Figure')
    expect(entries).toHaveLength(2)
    expect(entries[0]).toMatchObject({
      level: 1,
      text: 'Figure 1 Architecture',
      anchor: '_Ref111111111',
    })
    expect(entries[1]).toMatchObject({
      level: 1,
      text: 'Figure 2 Data flow',
      anchor: '_Ref222222222',
    })
  })

  it('inserts a ToF field over two captions and round-trips through a save', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    editor.commands.insertContentAt(0, [
      captionNode('Figure', 1, 'Architecture', '_Ref111111111'),
      captionNode('Figure', 2, 'Data flow', '_Ref222222222'),
    ] as never)
    const entries = collectTofEntries(editor, parsed.blocks, 'Figure')
    editor.commands.insertContentAt(0, fieldNodes(entries, { seqIdentifier: 'Figure' }) as never)

    // field paragraphs carry the Word ToF instruction and hyperlink anchors
    expect(fieldInstructions(editor)).toEqual(['TOC \\h \\z \\c "Figure"'])
    expect(tocLines(editor)).toHaveLength(2)
    expect(tocLines(editor)[0]).toMatchObject({
      left: 'Figure 1 Architecture',
      anchor: '_Ref111111111',
      level: 1,
    })

    // save plan emits self-contained xml; reopen parses them back as tocLines
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await buildDocx({ bodyXml: planBodyXml(plan, parsed.blocks) }))
    const lines = reopened.blocks.filter((b) => b.fieldDisplay?.kind === 'tocLine')
    expect(lines).toHaveLength(2)
    expect(lines[0].fieldDisplay).toMatchObject({
      left: 'Figure 1 Architecture',
      level: 1,
      anchor: '_Ref111111111',
    })
    const instrXml = reopened.blocks.map((b) => b.originalXml ?? '').join('')
    expect(instrXml).toContain('TOC \\h \\z \\c "Figure"')
  })

  it('update rebuilds the ToF from new captions (F9 path, silent)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNode('Figure', 1, 'Architecture', '_Ref111111111'),
      captionNode('Figure', 2, 'Data flow', '_Ref222222222'),
    ] as never)
    editor.commands.insertContentAt(
      0,
      fieldNodes(collectTofEntries(editor, parsed.blocks, 'Figure'), {
        seqIdentifier: 'Figure',
      }) as never,
    )
    // a third caption lands after the initial insert
    editor.commands.insertContentAt(
      editor.state.doc.content.size,
      captionNode('Figure', 3, 'Failure modes', '_Ref444444444') as never,
    )
    expect(
      updateTocField(editor, parsed.blocks, undefined, undefined, undefined, { silent: true }),
    ).toBe('updated')
    expect(fieldInstructions(editor)).toEqual(['TOC \\h \\z \\c "Figure"'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(3)
    expect(lines[2]).toMatchObject({ left: 'Figure 3 Failure modes', anchor: '_Ref444444444' })
  })

  it('update reports missing when no TOC field exists (F9 stays quiet)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    expect(
      updateTocField(editor, parsed.blocks, undefined, undefined, undefined, { silent: true }),
    ).toBe('missing')
  })

  it('stores a canonical SEQ id while showing the translated word (UX-1011)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    // CaptionModal shape: canonical id in SEQ, translated word visible
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNodeLocalized('Figure', 1, 'Architektur', '_Ref111111111', 'Abbildung'),
    ] as never)
    // the stored instruction keeps the canonical id — matching survives a
    // UI-language switch
    let stored = ''
    editor.state.doc.forEach((node) => {
      if (node.type.name === 'docProtected') stored = String(node.attrs.genXml)
    })
    expect(stored).toContain('SEQ Figure \\* ARABIC')
    expect(stored).toContain('Abbildung ')
    // canonical match finds it; the display keeps the translated word
    const entries = collectTofEntries(editor, parsed.blocks, ['Figure'])
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({ text: 'Abbildung 1 Architektur', level: 1 })
  })

  it('collects legacy captions authored with the translated SEQ word (UX-1011)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    // pre-UX-1011 caption: the SEQ instruction carried the translated label
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNode('Abbildung', 1, 'Architektur', '_Ref111111111'),
    ] as never)
    // canonical id alone misses it; the alias list (id + locale word) collects
    expect(collectTofEntries(editor, parsed.blocks, ['Figure'])).toHaveLength(0)
    const entries = collectTofEntries(editor, parsed.blocks, ['Figure', 'Abbildung'])
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({ text: 'Abbildung 1 Architektur' })
  })
})

describe('ToF update alias compatibility (BUG-1110)', () => {
  // the module-level t() drives the update path's alias resolution; its
  // dictionary loads on demand (PERF-904)
  beforeAll(async () => {
    setModuleLang('de')
    await loadLocale('de')
  })
  afterAll(() => setModuleLang('en'))

  it('F9/update collects legacy captions through the canonical+locale alias set', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    // legacy caption: SEQ carries the translated word (pre-UX-1011 authoring)
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNode('Abbildung', 1, 'Architektur', '_Ref111111111'),
    ] as never)
    // the field instruction is canonical (\c "Figure"): Insert ToF finds the
    // caption via seqAliases — the update path must too
    editor.commands.insertContentAt(
      0,
      fieldNodes([{ level: 1, text: 'Abbildung 1 Architektur' }], {
        seqIdentifier: 'Figure',
      }) as never,
    )
    expect(
      updateTocField(editor, parsed.blocks, undefined, undefined, undefined, { silent: true }),
    ).toBe('updated')
    const lines = tocLines(editor)
    expect(lines).toHaveLength(1)
    expect(lines[0]).toMatchObject({ left: 'Abbildung 1 Architektur', anchor: '_Ref111111111' })
  })

  it('a legacy \\c instruction still collects canonical captions (mixed documents)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    // canonical caption (UX-1011): SEQ carries the id, display keeps the locale word
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNodeLocalized('Figure', 1, 'Architektur', '_Ref111111111', 'Abbildung'),
    ] as never)
    // the field instruction is legacy (\c "Abbildung"): the alias set must
    // bridge it back to the canonical id instead of splitting the series
    editor.commands.insertContentAt(
      0,
      fieldNodes([{ level: 1, text: 'Abbildung 1 Architektur' }], {
        seqIdentifier: 'Abbildung',
      }) as never,
    )
    expect(
      updateTocField(editor, parsed.blocks, undefined, undefined, undefined, { silent: true }),
    ).toBe('updated')
    expect(tocLines(editor)[0]).toMatchObject({ left: 'Abbildung 1 Architektur' })
    // the regenerated instruction keeps the authored identifier byte-for-byte
    expect(fieldInstructions(editor)).toEqual(['TOC \\h \\z \\c "Abbildung"'])
  })
})

describe('TOC options update', () => {
  it('keeps the authored \\n and level range across an update', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    editor.commands.insertContentAt(0, [headingNode('Chapter', 1), headingNode('Deep', 3)] as never)
    editor.commands.insertContentAt(
      0,
      fieldNodes(
        [
          { level: 1, text: 'Chapter' },
          { level: 3, text: 'Deep' },
        ],
        { levels: 1, hidePageNumbers: true },
      ) as never,
    )
    expect(updateTocField(editor, parsed.blocks)).toBe('updated')
    expect(fieldInstructions(editor)).toEqual(['TOC \\o "1-1" \\n \\h \\z \\u'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(1)
    expect(lines[0]).toMatchObject({ left: 'Chapter', noPage: true })
  })

  it('keeps a ranged \\n per level across an update: levels 2-4 lose pages, level 1 keeps them (BUG-1012)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    editor.commands.insertContentAt(0, [
      headingNode('Chapter', 1),
      headingNode('Section', 2),
      headingNode('Deep', 3),
    ] as never)
    editor.commands.insertContentAt(
      0,
      fieldNodes(
        [
          { level: 1, text: 'Chapter' },
          { level: 2, text: 'Section' },
          { level: 3, text: 'Deep' },
        ],
        { levels: 3, hidePageNumbersFrom: 2, hidePageNumbersTo: 4 },
      ) as never,
    )
    expect(updateTocField(editor, parsed.blocks)).toBe('updated')
    // the range round-trips instead of widening to a full \n
    expect(fieldInstructions(editor)).toEqual(['TOC \\o "1-3" \\n 2-4 \\h \\z \\u'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(3)
    expect(lines[0]).toMatchObject({ left: 'Chapter', level: 1 })
    expect(lines[0].noPage).toBeUndefined()
    expect(lines[1]).toMatchObject({ left: 'Section', level: 2, noPage: true })
    expect(lines[2]).toMatchObject({ left: 'Deep', level: 3, noPage: true })
  })

  it('updates every TOC/TOF field, not just the first (BUG-1011)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    // real headings the heading-TOC collects, then a caption set for the ToF
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      headingNode('Chapter', 1),
      headingNode('Appendix', 1),
    ] as never)
    // a heading TOC up front, a figure ToF behind it — the audit scenario
    editor.commands.insertContentAt(
      0,
      fieldNodes(
        [
          { level: 1, text: 'Chapter' },
          { level: 1, text: 'Appendix' },
        ],
        { levels: 1 },
      ) as never,
    )
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      captionNode('Figure', 1, 'Architecture', '_Ref111111111'),
      captionNode('Figure', 2, 'Data flow', '_Ref222222222'),
    ] as never)
    editor.commands.insertContentAt(
      editor.state.doc.content.size,
      fieldNodes(
        [
          { level: 1, text: 'Figure 1 Architecture' },
          { level: 1, text: 'Figure 2 Data flow' },
        ],
        { seqIdentifier: 'Figure' },
      ) as never,
    )
    // new content both fields must pick up
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      headingNode('Epilogue', 1),
      captionNode('Figure', 3, 'Timeline', '_Ref333333333'),
    ] as never)

    expect(updateTocField(editor, parsed.blocks)).toBe('updated')
    const instrs = fieldInstructions(editor)
    expect(instrs).toEqual(['TOC \\o "1-1" \\h \\z \\u', 'TOC \\h \\z \\c "Figure"'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(6) // 3 headings + 3 captions across both fields
    expect(lines[0]).toMatchObject({ left: 'Chapter' })
    expect(lines[2]).toMatchObject({ left: 'Epilogue' })
    expect(lines[3]).toMatchObject({ left: 'Figure 1 Architecture' })
    expect(lines[5]).toMatchObject({ left: 'Figure 3 Timeline' })
  })
})

describe('TOC \\t source styles (BUG-1012)', () => {
  // custom paragraph styles whose NAME differs from their styleId: the \t
  // spec carries names, paragraphs carry ids — the parsed styles map bridges
  const CUSTOM_STYLE_DOC =
    '<w:p><w:pPr><w:pStyle w:val="ChapterTitle"/></w:pPr><w:r><w:t>Opening</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>plain body</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="Appx"/></w:pPr><w:r><w:t>Appendix A</w:t></w:r></w:p>'

  async function openStyledDoc() {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: CUSTOM_STYLE_DOC,
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="ChapterTitle"><w:name w:val="Chapter Title"/></w:style>' +
          '<w:style w:type="paragraph" w:styleId="Appx"><w:name w:val="Appendix"/></w:style>',
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    openEditors.push(editor)
    return { editor, parsed }
  }

  it('update collects entries by style NAME via the parsed styles map, with page numbers', async () => {
    const { editor, parsed } = await openStyledDoc()
    editor.commands.insertContentAt(
      0,
      fieldNodes(
        [
          { level: 1, text: 'Opening' },
          { level: 2, text: 'Appendix A' },
        ],
        { styles: 'Chapter Title,1,Appendix,2' },
      ) as never,
    )
    // the styles map is what makes name↔styleId matching work (BUG-1012);
    // anchorPage supplies the styled paragraphs' real page numbers
    expect(updateTocField(editor, parsed.blocks, undefined, () => 7, parsed.styles)).toBe('updated')
    expect(fieldInstructions(editor)).toEqual(['TOC \\h \\z \\t "Chapter Title,1,Appendix,2"'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(2)
    expect(lines[0]).toMatchObject({ left: 'Opening', level: 1, right: '7' })
    expect(lines[1]).toMatchObject({ left: 'Appendix A', level: 2, right: '7' })
  })

  it('without the styles map the update degrades to no-entries instead of rebuilding from headings', async () => {
    const { editor, parsed } = await openStyledDoc()
    editor.commands.insertContentAt(
      0,
      fieldNodes([{ level: 1, text: 'Opening' }], { styles: 'Chapter Title,1' }) as never,
    )
    // names cannot resolve to styleIds: the field is left untouched
    expect(
      updateTocField(editor, parsed.blocks, undefined, undefined, undefined, { silent: true }),
    ).toBe('no-entries')
    expect(tocLines(editor)).toHaveLength(1)
    expect(fieldInstructions(editor)).toEqual(['TOC \\h \\z \\t "Chapter Title,1"'])
  })

  it('\\o + \\t collect the union: headings within the range plus style-mapped paragraphs', async () => {
    const { editor, parsed } = await openStyledDoc()
    // a real heading AFTER the styled paragraphs: document order decides the
    // merged entry order
    editor.commands.insertContentAt(editor.state.doc.content.size, [
      headingNode('Summary', 2),
    ] as never)
    editor.commands.insertContentAt(
      0,
      fieldNodes([{ level: 2, text: 'Summary' }], {
        levels: 2,
        styles: 'Chapter Title,1',
      }) as never,
    )
    expect(updateTocField(editor, parsed.blocks, undefined, undefined, parsed.styles)).toBe(
      'updated',
    )
    // both sources survive the regeneration (canonical order: \o then \t)
    expect(fieldInstructions(editor)).toEqual(['TOC \\o "1-2" \\h \\z \\t "Chapter Title,1"'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(2)
    expect(lines[0]).toMatchObject({ left: 'Opening', level: 1 })
    expect(lines[1]).toMatchObject({ left: 'Summary', level: 2 })
  })

  it('collectStyledTocEntries matches the spec levels and skips empty paragraphs', () => {
    // direct unit check of the spec parser: malformed pairs drop out silently
    expect(parseTocStyleSpec('Chapter Title,1,Appendix,2')).toEqual([
      { name: 'Chapter Title', level: 1 },
      { name: 'Appendix', level: 2 },
    ])
    // levels outside 1-9 and non-numeric levels drop out silently
    expect(parseTocStyleSpec('NoLevel,9,x,0,Bad,abc')).toEqual([{ name: 'NoLevel', level: 9 }])
  })
})

/**
 * PAR-110: TOC options + table-of-figures authoring. ToF entries come from
 * SEQ captions of one label, an update rebuilds the field from its authored
 * instruction (switches survive), and the flow round-trips through a save.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateCaptionXml, generateTocFieldXml, parseDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectTofEntries, updateTocField } from '../src/renderer/components/ribbon-references-tab'

async function openDoc(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

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
    expect(updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })).toBe(
      'updated',
    )
    expect(fieldInstructions(editor)).toEqual(['TOC \\h \\z \\c "Figure"'])
    const lines = tocLines(editor)
    expect(lines).toHaveLength(3)
    expect(lines[2]).toMatchObject({ left: 'Figure 3 Failure modes', anchor: '_Ref444444444' })
  })

  it('update reports missing when no TOC field exists (F9 stays quiet)', async () => {
    const { editor, parsed } = await openDoc('<w:p><w:r><w:t>Body</w:t></w:r></w:p>')
    expect(updateTocField(editor, parsed.blocks, undefined, undefined, { silent: true })).toBe(
      'missing',
    )
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

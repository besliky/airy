import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections, saveDocx } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blockTexts,
  mergeCompareDocs,
  pmBlockTexts,
  type CompareStamp,
} from '../src/renderer/editor/compare'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  acceptAllRevisions,
  collectRevisions,
  rejectAllRevisions,
} from '../src/renderer/editor/revisions'

const para = (text: string, docxIndex: number | null = null): PmNode => ({
  type: 'docParagraph',
  attrs: { docxIndex },
  ...(text ? { content: [{ type: 'text', text }] } : {}),
})

const stamp: CompareStamp = { author: 'Compare', date: '2026-09-18T10:00:00Z' }

const markOf = (node: PmNode, type: string) => (node.marks ?? []).find((m) => m.type === type)

function createEditor(content: PmNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

/** two small contract revisions sharing a common spine */
const LEFT_TEXTS = ['Keep one', 'Remove me', 'Keep two', 'Change me please', 'Keep three']
const RIGHT_TEXTS = ['Keep one', 'Keep two', 'Change me now', 'Keep three', 'Brand new']

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const LEFT_XML = LEFT_TEXTS.map(p).join('')
const RIGHT_XML = RIGHT_TEXTS.map(p).join('')

describe('mergeCompareDocs (legal blackline merge)', () => {
  it('marks removed / added / changed paragraphs with ins / del revisions', () => {
    const left = LEFT_TEXTS.map((text, i) => para(text, i))
    const right = RIGHT_TEXTS.map((text, i) => para(text, i))
    const { content, summary } = mergeCompareDocs(left, right, stamp)
    // same, removed, same, changed, same, added
    expect(summary).toEqual({ added: 1, removed: 1, changed: 1 })
    expect(content).toHaveLength(6)
    // unchanged paragraph passes through untouched (byte-preserving anchor kept)
    expect(content[0]).toEqual(para('Keep one', 0))
    // paragraph missing from the compared file: struck through, anchor kept
    expect(content[1].attrs?.docxIndex).toBe(1)
    expect(markOf(content[1].content![0], 'del')?.attrs).toMatchObject({ author: 'Compare' })
    // changed paragraph: common stem kept, old tail struck, new tail underlined
    const changed = content[3]
    expect(changed.attrs?.docxIndex).toBe(3)
    expect(changed.content).toHaveLength(3)
    expect(changed.content![0]).toEqual({ type: 'text', text: 'Change me ' })
    expect(markOf(changed.content![1], 'del')).toBeTruthy()
    expect(changed.content![1].text).toBe('please')
    expect(markOf(changed.content![2], 'ins')).toBeTruthy()
    expect(changed.content![2].text).toBe('now')
    // paragraph coming from the compared file: underlined, foreign anchor stripped
    const added = content[5]
    expect(added.attrs?.docxIndex).toBeNull()
    expect(markOf(added.content![0], 'ins')?.attrs).toMatchObject({ author: 'Compare' })
    expect(added.content![0].text).toBe('Brand new')
  })

  it('records an empty removed paragraph as a block-level deletion', () => {
    const { content } = mergeCompareDocs(
      [para('Keep one'), para(''), para('Keep two')],
      [para('Keep one'), para('Keep two')],
      stamp,
    )
    expect(content[1].attrs?.blockRevision).toMatchObject({ kind: 'del', author: 'Compare' })
  })

  it('sanitizes foreign content: anchors, bookmarks and revision marks stripped', () => {
    const foreign = {
      type: 'docParagraph',
      attrs: { docxIndex: 7, bookmarks: ['target'], pPrChange: '{"author":"X"}' },
      content: [
        {
          type: 'text',
          text: 'added',
          marks: [
            { type: 'ins', attrs: { author: 'Other', date: '2020-01-01T00:00:00Z' } },
            { type: 'comment', attrs: { ids: 'c1' } },
          ],
        },
      ],
    } as PmNode
    const { content } = mergeCompareDocs([], [foreign], stamp)
    const attrs = content[0].attrs!
    expect(attrs.docxIndex).toBeNull()
    expect(attrs.bookmarks).toBeNull()
    expect(attrs.pPrChange).toBeNull()
    // the compared file's own revision mark is replaced by the compare insertion
    expect(markOf(content[0].content![0], 'ins')?.attrs).toMatchObject({ author: 'Compare' })
    expect((content[0].content![0].marks ?? []).map((m) => m.type)).not.toContain('comment')
  })

  it('swaps blocks wholesale when nothing survives the run-level diff', () => {
    const { content } = mergeCompareDocs([para('abc')], [para('xyz')], stamp)
    expect(content).toHaveLength(2)
    expect(markOf(content[0].content![0], 'del')).toBeTruthy()
    expect(markOf(content[1].content![0], 'ins')).toBeTruthy()
  })

  it('reports identical documents without creating revisions', () => {
    const left = LEFT_TEXTS.map((text, i) => para(text, i))
    const right = LEFT_TEXTS.map((text) => para(text))
    const { content, summary } = mergeCompareDocs(left, right, stamp)
    expect(summary).toEqual({ added: 0, removed: 0, changed: 0 })
    const editor = createEditor(content)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('degrades to index-paired blocks above the paragraph budget instead of freezing (BUG-913)', () => {
    // (2500+1)^2 = 6.25M cells > PARA_DIFF_BUDGET (4M): books used to build a
    // 25-100M-cell LCS matrix here and freeze the renderer
    const n = 2500
    const left = Array.from({ length: n }, (_, i) => para(`L${i}`))
    const right = Array.from({ length: n }, (_, i) => para(`R${i}`))
    const start = Date.now()
    const { content, summary, degraded } = mergeCompareDocs(left, right, stamp)
    expect(Date.now() - start).toBeLessThan(2000)
    expect(degraded).toBe(true)
    expect(summary.changed).toBe(n)
    // each changed pair is still run-level merged: old prefix struck, new underlined
    expect(content).toHaveLength(n)
    expect(markOf(content[0].content![0], 'del')).toBeTruthy()
    expect(markOf(content[0].content![1], 'ins')).toBeTruthy()
  })
})

describe('compare merge: accept / reject round-trip on real fixtures', () => {
  it('merges fixture documents, and accept-all yields the compared text', async () => {
    const leftBytes = await buildDocx({ bodyXml: LEFT_XML })
    const rightBytes = await buildDocx({ bodyXml: RIGHT_XML })
    const leftParsed = await parseDocx(leftBytes)
    const rightParsed = await parseDocx(rightBytes)
    const left = blocksToPmDoc(leftParsed.blocks, readSections(leftParsed)).content!
    const right = blocksToPmDoc(rightParsed.blocks, readSections(rightParsed)).content!
    expect(pmBlockTexts(left)).toEqual(LEFT_TEXTS)
    expect(pmBlockTexts(right)).toEqual(RIGHT_TEXTS)

    const { content } = mergeCompareDocs(left, right, stamp)
    const editor = createEditor(content)
    // the blackline is a live revision set: one ins + one del + the changed pair
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.length).toBeGreaterThanOrEqual(4)
    expect(ranges.every((r) => r.author === 'Compare')).toBe(true)

    acceptAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    const acceptedBlocks = editor.state.doc.content.content.map((b) => b.textContent)
    expect(acceptedBlocks).toEqual(RIGHT_TEXTS)
    editor.destroy()
  })

  it('reject-all restores the original text', async () => {
    const leftBytes = await buildDocx({ bodyXml: LEFT_XML })
    const rightBytes = await buildDocx({ bodyXml: RIGHT_XML })
    const leftParsed = await parseDocx(leftBytes)
    const rightParsed = await parseDocx(rightBytes)
    const left = blocksToPmDoc(leftParsed.blocks, readSections(leftParsed)).content!
    const right = blocksToPmDoc(rightParsed.blocks, readSections(rightParsed)).content!
    const { content } = mergeCompareDocs(left, right, stamp)
    const editor = createEditor(content)
    rejectAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    const rejectedBlocks = editor.state.doc.content.content.map((b) => b.textContent)
    expect(rejectedBlocks).toEqual(LEFT_TEXTS)
    editor.destroy()
  })

  it('accepts and rejects an empty deleted paragraph', async () => {
    const leftBytes = await buildDocx({ bodyXml: `${p('Keep one')}<w:p/>${p('Keep two')}` })
    const rightBytes = await buildDocx({ bodyXml: `${p('Keep one')}${p('Keep two')}` })
    const leftParsed = await parseDocx(leftBytes)
    const rightParsed = await parseDocx(rightBytes)
    const left = blocksToPmDoc(leftParsed.blocks, readSections(leftParsed)).content!
    const right = blocksToPmDoc(rightParsed.blocks, readSections(rightParsed)).content!
    const { content } = mergeCompareDocs(left, right, stamp)
    for (const [apply, expected] of [
      [acceptAllRevisions, ['Keep one', 'Keep two']],
      [rejectAllRevisions, ['Keep one', '', 'Keep two']],
    ] as const) {
      const editor = createEditor(content)
      apply(editor)
      expect(editor.state.doc.content.content.map((b) => b.textContent)).toEqual([...expected])
      editor.destroy()
    }
  })

  it('round-trips through save: blackline XML keeps w:ins/w:del, final text matches the compared file', async () => {
    const leftBytes = await buildDocx({ bodyXml: LEFT_XML })
    const rightBytes = await buildDocx({ bodyXml: RIGHT_XML })
    const leftParsed = await parseDocx(leftBytes)
    const rightParsed = await parseDocx(rightBytes)
    const left = blocksToPmDoc(leftParsed.blocks, readSections(leftParsed)).content!
    const right = blocksToPmDoc(rightParsed.blocks, readSections(rightParsed)).content!
    const { content } = mergeCompareDocs(left, right, stamp)

    // blackline save: the merged document with revisions still pending
    const blacklineEditor = createEditor(content)
    const blacklinePlan = pmDocToSavePlan(blacklineEditor.getJSON() as PmNode, leftParsed.blocks)
    const blacklineBytes = await saveDocx(leftParsed, blacklinePlan.saveBlocks)
    const blacklineXml = (await parseDocx(blacklineBytes)).internal.documentXml
    expect(blacklineXml).toContain('<w:ins ')
    expect(blacklineXml).toContain('<w:del ')
    blacklineEditor.destroy()

    // accepted save: text equals the compared document
    const acceptedEditor = createEditor(content)
    acceptAllRevisions(acceptedEditor)
    const acceptedPlan = pmDocToSavePlan(acceptedEditor.getJSON() as PmNode, leftParsed.blocks)
    const acceptedBytes = await saveDocx(leftParsed, acceptedPlan.saveBlocks)
    const acceptedParsed = await parseDocx(acceptedBytes)
    expect(blockTexts(acceptedParsed.blocks)).toEqual(RIGHT_TEXTS)
    acceptedEditor.destroy()

    // rejected save: text equals the original document
    const rejectedEditor = createEditor(content)
    rejectAllRevisions(rejectedEditor)
    const rejectedPlan = pmDocToSavePlan(rejectedEditor.getJSON() as PmNode, leftParsed.blocks)
    const rejectedBytes = await saveDocx(leftParsed, rejectedPlan.saveBlocks)
    const rejectedParsed = await parseDocx(rejectedBytes)
    expect(blockTexts(rejectedParsed.blocks)).toEqual(LEFT_TEXTS)
    rejectedEditor.destroy()
  })
})

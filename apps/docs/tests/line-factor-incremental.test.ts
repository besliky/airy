/**
 * The incremental line-factor decoration rebuild must produce exactly the same
 * decoration set as a full rebuild from scratch, for every kind of transaction
 * (typing, mark steps, block insertion/removal, multi-step replaces).
 * PERF-1639: the pre-fix path rebuilt the whole-document set on every
 * transaction, which collapsed large documents into an O(n^2) layout storm.
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

/** the lineFactorLive plugin's current decoration records (from/to + attrs) */
function lineFactorDecosOf(editor: Editor): string {
  const plugin = editor.state.plugins.find((p) =>
    (p as unknown as { key?: string }).key?.startsWith('lineFactorLive$'),
  )
  expect(plugin).toBeDefined()
  const state = plugin!.getState(editor.state) as {
    decos: { find: () => Array<{ from: number; to: number }> }
  }
  const rows = state.decos.find().map((d) => ({
    from: d.from,
    to: d.to,
    attrs: (d as unknown as { type?: { attrs?: unknown } }).type?.attrs ?? null,
  }))
  rows.sort((a, b) => a.from - b.from || a.to - b.to)
  return JSON.stringify(rows)
}

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

/** a fresh editor mounts through the full-rebuild init path: the equivalence oracle */
function fullRebuildSnapshot(json: unknown): string {
  const oracle = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: json as never,
  })
  editors.push(oracle)
  return lineFactorDecosOf(oracle)
}

const para = (text: string) => ({
  type: 'docParagraph',
  content: text ? [{ type: 'text', text }] : undefined,
})
const sizedPara = (text: string, sizeHalfPoints: number) => ({
  type: 'docParagraph',
  content: [
    {
      type: 'text',
      text,
      marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints } }],
    },
  ],
})
const makeEditor = (content: unknown) => {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })
  editors.push(editor)
  return editor
}

describe('incremental line-factor rebuild equals a full rebuild', () => {
  it('mark step over the whole run grows the strut and matches a full rebuild', () => {
    const editor = makeEditor({ type: 'doc', content: [sizedPara('sized', 24), para('tail')] })
    const before = lineFactorDecosOf(editor)
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28 })
    const after = lineFactorDecosOf(editor)
    expect(after).not.toEqual(before) // the strut grew 12pt -> 14pt
    expect(after).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
  })

  it('mark step on a partial range (strut unchanged) matches a full rebuild', () => {
    const editor = makeEditor({ type: 'doc', content: [sizedPara('sized', 24), para('tail')] })
    editor.commands.setTextSelection({ from: 1, to: 3 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28 })
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
  })

  it('unsetMark removes the strut and matches a full rebuild', () => {
    const editor = makeEditor({ type: 'doc', content: [sizedPara('sized', 24)] })
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.unsetMark('docTextStyle')
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
  })

  it('insertions, deletions and multi-step transactions stay equal to a full rebuild', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [sizedPara('alpha', 24), para('beta'), sizedPara('gamma delta', 28)],
    })
    // type into block 1
    editor.commands.setTextSelection({ from: 4, to: 4 })
    editor.commands.insertContent('X')
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
    // insert new paragraphs at the end (the phased-open stream path)
    const end = editor.state.doc.content.size
    editor.commands.insertContentAt(end, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'delta epsilon' }],
    })
    editor.commands.insertContentAt(editor.state.doc.content.size, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'CJK paragraph' }],
    })
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
    // delete a middle block (merge shifts every later position)
    const midFrom = editor.state.doc.child(1).nodeSize
    const midTo = midFrom + editor.state.doc.child(2).nodeSize
    editor.commands.deleteRange({ from: midFrom, to: midTo })
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
    // one tr with two replacements (append + mid-document insertion)
    const tr = editor.state.tr
    tr.insert(2, editor.schema.nodes.docParagraph.create(null, [editor.schema.text('lead')]))
    tr.insert(
      tr.doc.content.size,
      editor.schema.nodes.docParagraph.create(null, [editor.schema.text('zeta')]),
    )
    editor.view.dispatch(tr)
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
  })

  it('attribute-only step on a paragraph updates its decoration', () => {
    const editor = makeEditor({ type: 'doc', content: [para('plain'), para('rows')] })
    const tr = editor.state.tr
    tr.setNodeMarkup(0, undefined, { ...editor.state.doc.child(0).attrs, align: 'justify' })
    editor.view.dispatch(tr)
    expect(lineFactorDecosOf(editor)).toEqual(fullRebuildSnapshot(editor.state.doc.toJSON()))
  })
})

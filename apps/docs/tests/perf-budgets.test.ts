/**
 * PERF-1639 budgets and cheap-skip guards.
 *
 * Large documents must not pay O(document) work per transaction. These tests
 * pin the per-block counting helpers and the plugin guards that skip
 * whole-document walks while a document cannot contain the nodes they
 * decorate (deterministic: reference identity, no real timers).
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { countWords, countWordsInDoc } from '../src/renderer/word-count'
import { collectRevisions, countRevisionsInDoc } from '../src/renderer/editor/revisions'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
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

const para = (text: string) => ({
  type: 'docParagraph',
  content: text ? [{ type: 'text', text }] : undefined,
})
const ins = () => ({
  type: 'ins',
  attrs: { author: 'Alice', date: '2026-07-01T10:00:00Z' },
})

describe('countWordsInDoc', () => {
  it('sums per-block word counts; words never merge across blocks', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [para('first'), para('second'), para('中文 分词')],
    })
    // 'first' + 'second' are two Latin words; each CJK character counts one by
    // one (2 + 2 characters for the two CJK words worth of characters)
    expect(countWordsInDoc(editor.state.doc)).toBe(6)
    // per-block semantics: a trailing word of block A and a leading word of
    // block B stay two words (the old whole-doc textContent merged them)
    expect(countWordsInDoc(editor.state.doc)).not.toBe(countWords(editor.state.doc.textContent))
  })

  it('counts CJK characters one by one inside a block', () => {
    const editor = makeEditor({ type: 'doc', content: [para('中文')] })
    expect(countWordsInDoc(editor.state.doc)).toBe(2)
  })

  it('reuses the count of untouched blocks after an edit', () => {
    const editor = makeEditor({ type: 'doc', content: [para('alpha'), para('beta')] })
    const before = countWordsInDoc(editor.state.doc)
    editor.commands.insertContentAt(editor.state.doc.content.size, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'gamma' }],
    })
    expect(countWordsInDoc(editor.state.doc)).toBe(before + 1)
  })
})

describe('countRevisionsInDoc', () => {
  it('equals collectRevisions length on multi-block revision documents', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        para('KE '),
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'IN', marks: [ins()] }],
        },
        para('plain'),
        {
          type: 'docParagraph',
          attrs: { blockRevision: { kind: 'ins', author: 'Bob' } },
          content: [{ type: 'text', text: 'block' }],
        },
      ],
    })
    expect(countRevisionsInDoc(editor.state.doc)).toBe(collectRevisions(editor.state.doc).length)
    expect(countRevisionsInDoc(editor.state.doc)).toBeGreaterThan(0)
  })

  it('equals collectRevisions length when there are no revisions', () => {
    const editor = makeEditor({ type: 'doc', content: [para('a'), para('b')] })
    expect(countRevisionsInDoc(editor.state.doc)).toBe(0)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
  })
})

describe('listNumbering / anchorLine cheap-skip guards', () => {
  const plainDoc = {
    type: 'doc',
    content: Array.from({ length: 64 }, (_, i) => para(`plain ${i}`)),
  }

  it('a plain-paragraph insert leaves the listNumbering state untouched', () => {
    const editor = makeEditor(plainDoc)
    const stateBefore = listNumberingState(editor)
    editor.commands.insertContentAt(editor.state.doc.content.size, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'more plain' }],
    })
    expect(listNumberingState(editor)).toBe(stateBefore) // reference identity: no recompute
  })

  it('inserting a list item recomputes the listNumbering state', () => {
    const editor = makeEditor(plainDoc)
    const stateBefore = listNumberingState(editor)
    editor.commands.insertContentAt(editor.state.doc.content.size, {
      type: 'docListItem',
      attrs: { numId: null, ilvl: 0 },
      content: [{ type: 'text', text: 'item' }],
    })
    expect(listNumberingState(editor)).not.toBe(stateBefore)
  })

  it('a plain-paragraph insert leaves the anchorLine state untouched', () => {
    const editor = makeEditor(plainDoc)
    const stateBefore = anchorLineState(editor)
    editor.commands.insertContentAt(editor.state.doc.content.size, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'more plain' }],
    })
    expect(anchorLineState(editor)).toBe(stateBefore)
  })

  it('inserting a docProtected node recomputes the anchorLine state', () => {
    const editor = makeEditor(plainDoc)
    const stateBefore = anchorLineState(editor)
    const node = editor.schema.nodes.docProtected.create()
    const tr = editor.state.tr.insert(editor.state.doc.content.size, node)
    editor.view.dispatch(tr)
    expect(editor.state.doc.lastChild?.type.name).toBe('docProtected')
    expect(anchorLineState(editor)).not.toBe(stateBefore)
  })
})

function listNumberingState(editor: Editor): unknown {
  const plugin = editor.state.plugins.find((p) =>
    (p as unknown as { key?: string }).key?.startsWith('listNumbering$'),
  )
  expect(plugin).toBeDefined()
  return plugin!.getState(editor.state)
}

/**
 * The anchorLine plugin registers no PluginKey; identify it by its state shape
 * ({ decos, candidates }) — the only decoration plugin carrying that pair.
 */
function anchorLineState(editor: Editor): unknown {
  const plugin = editor.state.plugins.find((p) => {
    const s = p.getState(editor.state) as { candidates?: unknown } | undefined
    return !!s && typeof s === 'object' && 'candidates' in s && 'decos' in s
  })
  expect(plugin).toBeDefined()
  return plugin!.getState(editor.state)
}

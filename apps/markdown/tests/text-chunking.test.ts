import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Node as PMNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { giantTextChunkingKey } from '../src/renderer/editor/giantTextChunking'

// PERF-1700: giant single paragraphs are mirrored in the DOM as ~16KiB text
// nodes separated by <wbr> widget decorations (view-only). These tests pin
// the invariants that make that safe: no decorations on normal content, the
// document model and serialization are never touched, DOM edits rewrite only
// the piece containing the change (TextViewDesc reuse), and text survives a
// full edit/undo series byte for byte.

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

const CHUNK = 16_384
// 14 pieces (13 chunk boundaries); the 500-char tail keeps the last edit in
// these tests from landing exactly on a boundary
const LEN = CHUNK * 13 + 500

/** word salad with soft-wrap opportunities, exactly `len` chars */
function wordSalad(len: number): string {
  let s = ''
  while (s.length < len) s += 'lorem ipsum dolor sit amet consectetur adipiscing elit '
  return s.slice(0, len)
}

/** the <p> of the first paragraph in the editor DOM */
function firstParagraph(editor: Editor): HTMLElement {
  const p = editor.view.dom.querySelector('p')
  if (!p) throw new Error('paragraph element not found')
  return p as HTMLElement
}

/** direct child text nodes of the paragraph (the chunk pieces) */
function pieceTextNodes(p: HTMLElement): Text[] {
  return Array.from(p.childNodes).filter((n) => n.nodeType === 3) as Text[]
}

function wbrCount(p: HTMLElement): number {
  return Array.from(p.childNodes).filter((n) => n.nodeName === 'WBR').length
}

function parseBody(editor: Editor, body: string): PMNode {
  return editor.markdown.parse(body)
}

/** mount `body` and close the history group so undo covers edits only */
function setup(editor: Editor, body: string): void {
  editor.commands.setContent(editor.markdown.parse(body), {})
  editor.view.dispatch(closeHistory(editor.state.tr))
}

describe('giant paragraph DOM text chunking', () => {
  it('decorates only giant paragraphs (none on normal documents)', () => {
    const small = createEditor()
    small.commands.setContent(small.markdown.parse('A normal short paragraph.\n\nSecond one.'), {})
    const deco = giantTextChunkingKey.getState(small.state)
    expect(deco?.decos.find().length).toBe(0)
    expect(wbrCount(firstParagraph(small))).toBe(0)

    const giant = createEditor()
    const body = wordSalad(LEN)
    giant.commands.setContent(giant.markdown.parse(body), {})
    const decoGiant = giantTextChunkingKey.getState(giant.state)
    expect(decoGiant?.decos.find().length).toBe(13)
    expect(wbrCount(firstParagraph(giant))).toBe(13)
  })

  it('never changes the document model or serialization (view-only)', () => {
    const body = wordSalad(LEN)
    const chunked = createEditor()
    chunked.commands.setContent(chunked.markdown.parse(body), {})

    // a second editor holds the same parsed document; the chunking plugin
    // must not make the first doc differ from the plain parse in any way
    const plain = createEditor()
    plain.commands.setContent(parseBody(chunked, body), {})

    expect(chunked.getJSON()).toEqual(plain.getJSON())
    expect(chunked.getMarkdown()).toBe(plain.getMarkdown())
    // text survives round trip byte for byte
    expect(chunked.state.doc.textContent).toBe(body)
    expect(chunked.getMarkdown()).toBe(body)
  })

  it('a tail edit rewrites only the last DOM piece and reuses the other nodes', () => {
    const body = wordSalad(LEN)
    const editor = createEditor()
    setup(editor, body)
    const p = firstParagraph(editor)
    const before = pieceTextNodes(p).map((n) => n.nodeValue)
    expect(before.length).toBe(14)

    // mark an early piece to prove DOM nodes are reused, not recreated
    const marker = Symbol('kept')
    const marked = pieceTextNodes(p)[3] as unknown as Record<symbol, unknown>
    marked[marker] = true

    editor.commands.insertContentAt(1 + body.length, 'x')

    const after = pieceTextNodes(p).map((n) => n.nodeValue)
    expect(after.length).toBe(14)
    expect((pieceTextNodes(p)[3] as unknown as Record<symbol, unknown>)[marker]).toBe(true)
    for (let i = 0; i < 13; i++) expect(after[i]).toBe(before[i])
    expect(after[13]).toBe(before[13] + 'x')
    expect(wbrCount(p)).toBe(13)
    // DOM text matches the model bit for bit
    expect(p.textContent).toBe(body + 'x')
    expect(editor.state.doc.textContent).toBe(body + 'x')
  })

  it('a mid-paragraph edit changes exactly one piece; undo restores byte for byte', () => {
    const body = wordSalad(LEN)
    const editor = createEditor()
    setup(editor, body)
    const p = firstParagraph(editor)
    const before = pieceTextNodes(p).map((n) => n.nodeValue)
    const jsonBefore = editor.getJSON()

    // well inside piece 5, away from chunk boundaries
    const at = CHUNK * 5 + 7
    editor.commands.insertContentAt(1 + at, 'EDIT')

    const after = pieceTextNodes(p).map((n) => n.nodeValue)
    const changed = after.filter((v, i) => v !== before[i]).length
    expect(changed).toBe(1)
    expect(wbrCount(p)).toBe(13)

    const expected = body.slice(0, at) + 'EDIT' + body.slice(at)
    expect(p.textContent).toBe(expected)
    expect(editor.state.doc.textContent).toBe(expected)

    expect(editor.can().undo()).toBe(true)
    editor.commands.undo()
    expect(p.textContent).toBe(body)
    expect(pieceTextNodes(p).map((n) => n.nodeValue)).toEqual(before)
    expect(editor.getJSON()).toEqual(jsonBefore)
  })

  it('re-anchors the grid when a piece outgrows REANCHOR_MAX, serialization intact', () => {
    const body = wordSalad(LEN)
    const editor = createEditor()
    setup(editor, body)
    const p = firstParagraph(editor)

    // push the tail piece far past 2*CHUNK so mapping can no longer keep the
    // coverage bounded — boundaries must be re-gridded from the start
    const tail = 'z'.repeat(CHUNK * 2 + 1000)
    editor.commands.insertContentAt(1 + body.length, tail)

    const expected = body + tail
    expect(p.textContent).toBe(expected)
    expect(editor.state.doc.textContent).toBe(expected)

    const lengths = pieceTextNodes(p).map((n) => n.nodeValue!.length)
    const expectedChunks = Math.floor(expected.length / CHUNK)
    expect(wbrCount(p)).toBe(expectedChunks)
    expect(lengths.length).toBe(expectedChunks + 1)
    // re-gridded: every non-final piece is exactly CHUNK long again
    for (let i = 0; i < lengths.length - 1; i++) expect(lengths[i]).toBe(CHUNK)

    // the model still serializes exactly like the plain parse of the same text
    const plain = createEditor()
    plain.commands.setContent(plain.markdown.parse(expected), {})
    expect(editor.getJSON()).toEqual(plain.getJSON())
    expect(editor.getMarkdown()).toBe(plain.getMarkdown())
  })

  it('a piece keeps growing while the caret stays inside it (no boundary drift)', () => {
    const body = wordSalad(LEN)
    const editor = createEditor()
    setup(editor, body)
    const p = firstParagraph(editor)
    const before = pieceTextNodes(p).map((n) => n.nodeValue)

    // several small edits inside piece 5: earlier and later pieces keep their
    // byte-identical DOM text; only the edited piece is rewritten
    let expected = body
    for (let i = 0; i < 5; i++) {
      const at = CHUNK * 5 + 7 + i * 10
      editor.commands.insertContentAt(1 + at, 'ab')
      expected = expected.slice(0, at) + 'ab' + expected.slice(at)
    }
    expect(p.textContent).toBe(expected)
    const after = pieceTextNodes(p).map((n) => n.nodeValue)
    expect(after.length).toBe(before.length)
    let changed = 0
    for (let i = 0; i < after.length; i++) if (after[i] !== before[i]) changed++
    expect(changed).toBe(1)
    expect(wbrCount(p)).toBe(13)
  })
})

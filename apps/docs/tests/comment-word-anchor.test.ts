import { describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { wordRangeAtCaret } from '../src/renderer/editor/comments'
import { startNewComment, type ReviewContext } from '../src/renderer/review-actions'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

/** put a collapsed caret at doc position `pos` */
function setCaret(editor: Editor, pos: number): void {
  const tr = editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos))
  editor.view.dispatch(tr)
}

const textAt = (editor: Editor, range: { from: number; to: number }): string =>
  editor.state.doc.textBetween(range.from, range.to)

/** minimal ReviewContext: only the editor and the composer/status callbacks matter here */
function reviewContext(editor: Editor): ReviewContext {
  return {
    editor,
    doc: null,
    dirtyRef: { current: false },
    setStatus: vi.fn(),
    notePrompt: null,
    setNotePrompt: vi.fn(),
    footnotes: [],
    endnotes: [],
    setFootnotes: vi.fn(),
    setEndnotes: vi.fn(),
    setNotesDirty: vi.fn(),
    comments: [],
    setComments: vi.fn(),
    setCommentsDirty: vi.fn(),
    setCommentComposing: vi.fn(),
    setShowComments: vi.fn(),
    setInkAnnotations: vi.fn(),
    setInksDirty: vi.fn(),
    setCompareResult: vi.fn(),
    setRevisionDisplay: vi.fn(),
  }
}

describe('wordRangeAtCaret', () => {
  // "alpha beta gamma": paragraph content starts at doc position 1
  const editor = () => createEditor([para(text('alpha beta gamma'))])

  it('picks the word around a caret inside it', () => {
    const ed = editor()
    setCaret(ed, 9) // between "be|ta"
    const range = wordRangeAtCaret(ed)
    expect(range && textAt(ed, range)).toBe('beta')
    ed.destroy()
  })

  it('picks the word the caret sits at the start of', () => {
    const ed = editor()
    setCaret(ed, 7) // "|beta"
    const range = wordRangeAtCaret(ed)
    expect(range && textAt(ed, range)).toBe('beta')
    ed.destroy()
  })

  it('picks the word just before a caret at its end', () => {
    const ed = editor()
    setCaret(ed, 6) // "alpha|"
    const range = wordRangeAtCaret(ed)
    expect(range && textAt(ed, range)).toBe('alpha')
    ed.destroy()
  })

  it('returns null with the caret in an empty paragraph', () => {
    const ed = createEditor([{ type: 'docParagraph', attrs: { docxIndex: null } }])
    setCaret(ed, 1)
    expect(wordRangeAtCaret(ed)).toBeNull()
    ed.destroy()
  })

  it('returns null for a non-empty selection', () => {
    const ed = editor()
    const tr = ed.state.tr.setSelection(TextSelection.create(ed.state.doc, 1, 6))
    ed.view.dispatch(tr)
    expect(wordRangeAtCaret(ed)).toBeNull()
    ed.destroy()
  })

  it('finds a word in CJK text', () => {
    const ed = createEditor([para(text('\u8fd9\u662f\u4e00\u4e2a\u6d4b\u8bd5'))])
    setCaret(ed, 4)
    const range = wordRangeAtCaret(ed)
    expect(range).not.toBeNull()
    // the segmenter's exact CJK word boundaries may vary; the caret must sit inside the range
    expect(range!.from).toBeLessThanOrEqual(4)
    expect(range!.to).toBeGreaterThanOrEqual(4)
    expect(range!.to).toBeGreaterThan(range!.from)
    ed.destroy()
  })

  it('anchors on the preceding word after trailing punctuation is skipped over', () => {
    const ed = createEditor([para(text('done. next'))])
    setCaret(ed, 5) // "done|."
    const range = wordRangeAtCaret(ed)
    expect(range && textAt(ed, range)).toBe('done')
    ed.destroy()
  })

  // UX-1711: a caret resting in whitespace between words must still anchor a
  // comment (Word never refuses) — the nearest word wins, next-word as fallback
  it('falls forward to the next word from a leading gap', () => {
    const ed = createEditor([para(text(' alpha'))])
    setCaret(ed, 1) // "| alpha" — on the space before the word
    const range = wordRangeAtCaret(ed)
    expect(range && textAt(ed, range)).toBe('alpha')
    ed.destroy()
  })

  it('falls forward to the next word from the middle of a whitespace run', () => {
    const ed = createEditor([para(text('a   b'))])
    setCaret(ed, 3) // "a  |b" — two spaces in, no word ends there
    const range = wordRangeAtCaret(ed)
    expect(range && textAt(ed, range)).toBe('b')
    ed.destroy()
  })

  it('still returns null in a paragraph without any word', () => {
    const ed = createEditor([para(text(' ... '))])
    setCaret(ed, 2)
    expect(wordRangeAtCaret(ed)).toBeNull()
    ed.destroy()
  })
})

// UX-1711: startNewComment must anchor on the word at/nearest a collapsed
// caret instead of refusing (Word semantics), and only refuse in a wordless block
describe('startNewComment caret anchoring', () => {
  it('selects the next word when the caret rests in a gap and opens the composer', () => {
    const ed = createEditor([para(text('a   b'))])
    setCaret(ed, 3) // "a  |b"
    const ctx = reviewContext(ed)
    startNewComment(ctx)
    const sel = ed.state.selection
    expect(sel.empty).toBe(false)
    expect(textAt(ed, { from: sel.from, to: sel.to })).toBe('b')
    expect(ctx.setShowComments).toHaveBeenCalledWith(true)
    expect(ctx.setCommentComposing).toHaveBeenCalledWith(true)
    ed.destroy()
  })

  it('refuses with the status hint only in a wordless paragraph', () => {
    const ed = createEditor([para(text(' ... '))])
    setCaret(ed, 2)
    const ctx = reviewContext(ed)
    startNewComment(ctx)
    // t() falls back to the raw key with no dictionary loaded (test env)
    expect(ctx.setStatus).toHaveBeenCalledWith('appSelectTextToComment')
    expect(ctx.setCommentComposing).not.toHaveBeenCalled()
    ed.destroy()
  })
})

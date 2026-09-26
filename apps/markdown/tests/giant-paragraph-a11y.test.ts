import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { giantParagraphA11yKey } from '../src/renderer/editor/giantParagraphA11y'

// PERF-1700b: giant paragraphs are hidden from the accessibility tree
// (aria-hidden) because Chromium's AX tree update otherwise taxes every
// keystroke in the document by up to ~2s (measured: the whole user-visible
// stall on a 4MB one-liner is LocalFrameView::RunAccessibilitySteps). These
// tests pin the policy invariants: normal documents are untouched, the
// attribute is applied through PM node decorations (so it survives
// prosemirror-view's attribute syncing), the visually-hidden summary widget
// sits in front of the paragraph and reflects its size, and nothing here
// changes the document model or serialization.

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

const LEN = 100_100

/** word salad with soft-wrap opportunities, exactly `len` chars */
function wordSalad(len: number): string {
  let s = ''
  while (s.length < len) s += 'lorem ipsum dolor sit amet consectetur adipiscing elit '
  return s.slice(0, len)
}

function firstParagraph(editor: Editor): HTMLElement {
  const p = editor.view.dom.querySelector('p')
  if (!p) throw new Error('paragraph element not found')
  return p as HTMLElement
}

function summaryWidget(editor: Editor): HTMLElement | null {
  return editor.view.dom.querySelector('.md-giant-para-summary')
}

function setup(editor: Editor, body: string): void {
  editor.commands.setContent(editor.markdown.parse(body), {})
  editor.view.dispatch(closeHistory(editor.state.tr))
}

describe('giant paragraph accessibility policy', () => {
  it('touches only giant paragraphs (no aria-hidden, no summary on normal docs)', () => {
    const small = createEditor()
    small.commands.setContent(small.markdown.parse('A normal short paragraph.\n\nSecond one.'), {})
    expect(giantParagraphA11yKey.getState(small.state)?.decos.find().length).toBe(0)
    expect(firstParagraph(small).getAttribute('aria-hidden')).toBe(null)
    expect(summaryWidget(small)).toBe(null)

    const giant = createEditor()
    giant.commands.setContent(giant.markdown.parse(wordSalad(LEN)), {})
    expect(firstParagraph(giant).getAttribute('aria-hidden')).toBe('true')
    expect(summaryWidget(giant)).not.toBe(null)
  })

  it('keeps aria-hidden across edits (PM-managed attribute, not ad-hoc DOM writes)', () => {
    const body = wordSalad(LEN)
    const editor = createEditor()
    setup(editor, body)
    expect(firstParagraph(editor).getAttribute('aria-hidden')).toBe('true')

    editor.commands.insertContentAt(1 + body.length, 'x')
    expect(firstParagraph(editor).getAttribute('aria-hidden')).toBe('true')
    expect(editor.state.doc.textContent).toBe(body + 'x')
  })

  it('summary widget states the paragraph size and tracks its growth', () => {
    const body = wordSalad(LEN)
    const editor = createEditor()
    setup(editor, body)

    const before = summaryWidget(editor)
    expect(before?.textContent).toContain(body.length.toLocaleString('en-US'))

    editor.commands.insertContentAt(1 + body.length, 'xyz')
    const after = summaryWidget(editor)
    expect(after?.textContent).toContain((body.length + 3).toLocaleString('en-US'))
  })

  it('never changes the document model or serialization (view-only)', () => {
    const body = wordSalad(LEN)
    const policy = createEditor()
    setup(policy, body)

    const plain = createEditor()
    plain.commands.setContent(plain.markdown.parse(body), {})

    expect(policy.getJSON()).toEqual(plain.getJSON())
    expect(policy.getMarkdown()).toBe(plain.getMarkdown())
    expect(policy.state.doc.textContent).toBe(body)
  })
})

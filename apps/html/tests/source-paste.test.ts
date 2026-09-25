import { afterEach, describe, expect, it } from 'vitest'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import {
  consumePlainPasteGesture,
  insertTextAtSelection,
  notePlainPasteGesture,
  sourcePaste,
} from '../src/renderer/source/source-paste'

// UX-1705: paste-flavor routing in the HTML source editor. A text/html
// clipboard flavor inserts sanitized markup (previously it was ignored and
// Word fragments landed as bare text/plain); Mod+Shift+V forces the plain
// flavor; a clipboard without the flavor keeps CodeMirror's default path.

const editors: EditorView[] = []
afterEach(() => {
  for (const view of editors.splice(0)) view.destroy()
  // disarm a gesture armed by a preceding test
  consumePlainPasteGesture()
})

function setup(doc: string): EditorView {
  const view = new EditorView({
    state: EditorState.create({ doc }),
    parent: document.body,
  })
  editors.push(view)
  return view
}

function clipboard(flavors: Record<string, string>): ClipboardEvent {
  return {
    clipboardData: { getData: (type: string) => flavors[type] ?? '' },
  } as unknown as ClipboardEvent
}

function keyboard(init: {
  key: string
  ctrlKey?: boolean
  metaKey?: boolean
  shiftKey?: boolean
  altKey?: boolean
}): KeyboardEvent {
  return init as unknown as KeyboardEvent
}

const WORD_HTML =
  '<!--StartFragment--><p class=MsoNormal style="font-family:x">Say <b>bold</b> fine</p>' +
  '<!--[if gte mso 9]><xml><w:WordDocument/></xml><![endif]--><!--EndFragment-->'

describe('text/html flavor inserts sanitized markup', () => {
  it('Word fragments land as clean markup, junk cut', () => {
    const view = setup('<p>doc</p>')
    view.dispatch({ selection: { anchor: view.state.doc.length } })
    const handled = sourcePaste(
      view,
      clipboard({ 'text/html': WORD_HTML, 'text/plain': 'Say bold fine' }),
    )
    expect(handled).toBe(true)
    const doc = view.state.doc.toString()
    expect(doc).toContain('Say <b>bold</b> fine')
    expect(doc).not.toContain('mso')
    expect(doc).not.toContain('<span')
    expect(doc).not.toContain('<!--')
    expect(doc).not.toContain('style=')
    // cursor sits behind the insertion
    expect(view.state.selection.main.head).toBe(view.state.doc.length)
  })

  it('replaces the selection when pasting over a range', () => {
    const view = setup('<h1>old</h1>')
    view.dispatch({ selection: { anchor: 0, head: view.state.doc.length } })
    sourcePaste(view, clipboard({ 'text/html': '<p>new</p>' }))
    expect(view.state.doc.toString()).toBe('<p>new</p>')
  })

  it('sanitizer-empty html falls back to the plain flavor', () => {
    const view = setup('<p>x</p>')
    const handled = sourcePaste(
      view,
      clipboard({ 'text/html': '<style>.x{}</style>', 'text/plain': 'fallback' }),
    )
    expect(handled).toBe(true)
    // the default selection sits at the document start
    expect(view.state.doc.toString()).toBe('fallback<p>x</p>')
  })
})

describe('plain paths', () => {
  it('plain-only clipboard is not intercepted (CodeMirror default stays)', () => {
    const view = setup('<p>doc</p>')
    expect(sourcePaste(view, clipboard({ 'text/plain': 'just text' }))).toBe(false)
    expect(view.state.doc.toString()).toBe('<p>doc</p>')
  })

  it('Mod+Shift+V inserts the literal plain text, markup untouched', () => {
    const view = setup('<p>doc</p>')
    view.dispatch({ selection: { anchor: view.state.doc.length } })
    notePlainPasteGesture(keyboard({ key: 'v', ctrlKey: true, shiftKey: true }))
    const handled = sourcePaste(
      view,
      clipboard({ 'text/html': '<p><b>x</b></p>', 'text/plain': '<b>x</b>' }),
    )
    expect(handled).toBe(true)
    expect(view.state.doc.toString()).toBe('<p>doc</p><b>x</b>')
    // the gesture is consumed
    expect(consumePlainPasteGesture()).toBe(false)
  })

  it('an armed gesture with an empty plain flavor falls through to the default', () => {
    const view = setup('<p>doc</p>')
    notePlainPasteGesture(keyboard({ key: 'V', ctrlKey: true, shiftKey: true }))
    expect(sourcePaste(view, clipboard({ 'text/html': '<p>x</p>' }))).toBe(false)
  })

  it('plain Ctrl+V does not arm the gesture', () => {
    notePlainPasteGesture(keyboard({ key: 'v', ctrlKey: true }))
    expect(consumePlainPasteGesture()).toBe(false)
  })
})

describe('insertTextAtSelection', () => {
  it('inserts verbatim and moves the caret behind the text', () => {
    const view = setup('<p>a</p>')
    view.dispatch({ selection: { anchor: view.state.doc.length } })
    insertTextAtSelection(view, '<i>b</i>')
    expect(view.state.doc.toString()).toBe('<p>a</p><i>b</i>')
    expect(view.state.selection.main.head).toBe(view.state.doc.length)
  })
})

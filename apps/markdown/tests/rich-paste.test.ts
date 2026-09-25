import { afterEach, describe, expect, it } from 'vitest'
import type { Editor } from '@tiptap/core'
import { Slice } from '@tiptap/pm/model'
import { linkDiagnosticsPluginKey } from '../src/renderer/editor/linkDiagnostics'
import {
  consumePlainPasteGesture,
  insertPlainText,
  notePlainPasteGesture,
} from '../src/renderer/editor/richPaste'

// UX-1705: paste-flavor routing in the markdown editor. A text/html clipboard
// flavor is sanitized and parsed through the schema (rawHtml policy); a plain
// paste keeps the historic path; Mod+Shift+V forces the plain flavor; the
// ProseMirror-internal slice (data-pm-slice) must keep the default handler.

// Undestroyed views leave DOMObserver flush timers that fire after jsdom
// teardown ("document is not defined" unhandled error) — destroy everything.
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
  // disarm a gesture armed by a preceding test
  consumePlainPasteGesture()
})

async function createEditor(): Promise<Editor> {
  const { Editor } = await import('@tiptap/core')
  const { buildExtensions } = await import('../src/renderer/editor/extensions')
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

/** synthetic clipboard: flavors map + files, mirroring the DataTransfer surface the handlers read */
function clipboard(flavors: Record<string, string>, files: File[] = []): DataTransfer {
  return { files, getData: (type: string) => flavors[type] ?? '' } as unknown as DataTransfer
}

/** run the editor's handlePaste props exactly like ProseMirror: first truthy wins */
function paste(editor: Editor, data: DataTransfer): boolean {
  const event = { clipboardData: data } as unknown as ClipboardEvent
  let handled = false
  editor.view.someProp('handlePaste', (fn) => {
    if (handled) return
    handled = fn(editor.view, event, Slice.empty) === true
  })
  return handled
}

function plainText(editor: Editor): string {
  const doc = editor.state.doc
  return doc.textBetween(0, doc.content.size, '\n', (node) =>
    node.type.name === 'hardBreak' ? '\n' : '',
  )
}

const WORD_HTML =
  '<!--StartFragment--><p class=MsoNormal style="font-family:x">Say <b>bold</b> fine</p>' +
  '<ul><li>one</li></ul><!--EndFragment-->'

describe('text/html flavor routes through the sanitizer', () => {
  it('inserts sanitized markup: formatting parses, mso junk never lands', async () => {
    const editor = await createEditor()
    const handled = paste(
      editor,
      clipboard({ 'text/html': WORD_HTML, 'text/plain': 'Say bold fine' }),
    )
    expect(handled).toBe(true)

    const md = editor.getMarkdown()
    // formatting survived as real GFM, not as a raw chip
    expect(md).toContain('**bold**')
    expect(md).toContain('- one')
    // the mso/comment wrappers and style attributes are gone
    expect(md).not.toContain('mso')
    expect(md).not.toContain('<span')
    expect(md).not.toContain('<!--')
    expect(plainText(editor)).toContain('Say bold fine')
  })

  it('markup the schema has no node for still lands as a rawHtml chip', async () => {
    const editor = await createEditor()
    paste(editor, clipboard({ 'text/html': '<p>hi <mark>marked</mark></p>' }))
    expect(editor.getMarkdown()).toContain('<mark>marked</mark>')
  })

  it('sanitizer-empty html falls back to the plain flavor', async () => {
    const editor = await createEditor()
    const handled = paste(
      editor,
      clipboard({ 'text/html': '<style>.x{}</style>', 'text/plain': 'fallback' }),
    )
    expect(handled).toBe(true)
    expect(plainText(editor)).toBe('fallback')
  })
})

describe('flavors and gestures that keep the default behavior', () => {
  it('plain-only clipboard is not intercepted (historic path)', async () => {
    const editor = await createEditor()
    const handled = paste(editor, clipboard({ 'text/plain': 'just text' }))
    expect(handled).toBe(false)
    expect(editor.state.doc.textContent).toBe('')
  })

  it('a ProseMirror-internal slice (data-pm-slice) is left to the exact-slice restore', async () => {
    const editor = await createEditor()
    const internal = '<div data-pm-slice="1 1 []"><p>internal</p></div>'
    const handled = paste(editor, clipboard({ 'text/html': internal, 'text/plain': 'internal' }))
    expect(handled).toBe(false)
  })

  it('pasting markup inside a code block stays plain (default handler)', async () => {
    const editor = await createEditor()
    editor.commands.setContent({ type: 'codeBlock', content: [{ type: 'text', text: 'x' }] }, {})
    editor.commands.setTextSelection(2)
    const handled = paste(editor, clipboard({ 'text/html': '<p><b>b</b></p>', 'text/plain': 'b' }))
    expect(handled).toBe(false)
    expect(editor.state.doc.firstChild!.type.name).toBe('codeBlock')
  })
})

describe('internal links survive paste (BUG-1739)', () => {
  // the round-3 audit vector: a Word fragment whose TOC links through anchors
  // and relative paths used to lose its hrefs in the sanitizer, so the links
  // died as plain text BEFORE link diagnostics (#198) could ever classify them
  const TOC_HTML =
    '<h1>Section Two</h1>' +
    '<a href="#section-two">internal jump</a>' +
    '<a href="#ghost-anchor">dead anchor</a>' +
    '<a href="docs/page.md">relative jump</a>'

  it('anchors and relative hrefs land as real links, not plain text', async () => {
    const editor = await createEditor()
    const handled = paste(editor, clipboard({ 'text/html': TOC_HTML, 'text/plain': 'text' }))
    expect(handled).toBe(true)
    const md = editor.getMarkdown()
    expect(md).toContain('[internal jump](#section-two)')
    expect(md).toContain('[dead anchor](#ghost-anchor)')
    expect(md).toContain('[relative jump](docs/page.md)')
    expect(md).toContain('# Section Two')
  })

  it('link diagnostics sees the pasted anchors: resolved one unmarked, ghost marked dead', async () => {
    const editor = await createEditor()
    paste(editor, clipboard({ 'text/html': TOC_HTML, 'text/plain': 'text' }))
    const state = linkDiagnosticsPluginKey.getState(editor.state)
    expect(state).toBeTruthy()
    const dead = (state?.set.find() ?? [])
      .map((d) => ({
        text: editor.state.doc.textBetween(d.from, d.to, ' ', ' '),
        cls: String(d.type.attrs?.class ?? ''),
      }))
      .filter((d) => d.cls.startsWith('md-link-'))
    // the pasted <h1>Section Two</h1> makes #section-two resolvable — exactly
    // the irony the audit called out; only the ghost anchor is honestly dead
    expect(dead).toEqual([{ text: 'dead anchor', cls: 'md-link-dead-anchor' }])
  })
})

describe('Mod+Shift+V forces the plain flavor', () => {
  it('keydown arms the gesture; the next paste inserts literal text only', async () => {
    const editor = await createEditor()
    notePlainPasteGesture({
      key: 'V',
      ctrlKey: true,
      metaKey: false,
      shiftKey: true,
      altKey: false,
    } as KeyboardEvent)
    const handled = paste(
      editor,
      clipboard({ 'text/html': WORD_HTML, 'text/plain': 'Say <b>bold</b> fine' }),
    )
    expect(handled).toBe(true)
    expect(plainText(editor)).toBe('Say <b>bold</b> fine')
    // no markup model was created
    expect(JSON.stringify(editor.state.doc.toJSON())).not.toContain('strong')
    // the gesture is consumed — a following paste routes normally
    expect(consumePlainPasteGesture()).toBe(false)
  })

  it('a Mod+Shift+V gesture without a matching paste expires', () => {
    notePlainPasteGesture({
      key: 'v',
      ctrlKey: true,
      metaKey: false,
      shiftKey: true,
      altKey: false,
    } as KeyboardEvent)
    expect(consumePlainPasteGesture()).toBe(true)
  })

  it('plain Ctrl+V does not arm the gesture', () => {
    notePlainPasteGesture({
      key: 'v',
      ctrlKey: true,
      metaKey: false,
      shiftKey: false,
      altKey: false,
    } as KeyboardEvent)
    expect(consumePlainPasteGesture()).toBe(false)
  })
})

describe('insertPlainText (paste-as-plain-text command)', () => {
  it('inserts text verbatim: newlines become hard breaks, no marks', async () => {
    const editor = await createEditor()
    insertPlainText(editor, 'alpha\nbeta')
    const json = JSON.stringify(editor.state.doc.toJSON())
    expect(json).toContain('"hardBreak"')
    // literal text only — no formatting model around it
    expect(json).not.toContain('strong')
    expect(plainText(editor)).toBe('alpha\nbeta')
  })

  it('normalizes CRLF and ignores empty input', async () => {
    const editor = await createEditor()
    insertPlainText(editor, 'a\r\nb\rc')
    expect(plainText(editor)).toBe('a\nb\nc')
    const before = JSON.stringify(editor.state.doc.toJSON())
    insertPlainText(editor, '')
    expect(JSON.stringify(editor.state.doc.toJSON())).toBe(before)
  })
})

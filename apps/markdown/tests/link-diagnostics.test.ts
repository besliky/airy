import { afterEach, describe, expect, it, vi } from 'vitest'
import { setToastEmitter, type ToastData } from '@airy-office/ui/toast-bus'
import {
  headingSlug,
  linkIssue,
  linkDiagnosticsPluginKey,
} from '../src/renderer/editor/linkDiagnostics'
import { strings } from '../src/renderer/i18n/strings'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: import('@tiptap/core').Editor[] = []
let toasts: ToastData[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
  setToastEmitter(null)
  toasts = []
  vi.restoreAllMocks()
})

async function newEditor(markdown: string) {
  const { Editor } = await import('@tiptap/core')
  const { buildExtensions } = await import('../src/renderer/editor/extensions')
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: markdown,
    contentType: 'markdown',
    // mount so the extension lifecycle hooks ('create') actually run, as in the app
    element: document.createElement('div'),
  })
  editors.push(editor)
  return editor
}

interface Found {
  text: string
  cls: string
}

/** Painted diagnostics of the editor, as {decorated text, decoration class} */
function diagnostics(editor: import('@tiptap/core').Editor): Found[] {
  const state = linkDiagnosticsPluginKey.getState(editor.state)
  expect(state).toBeTruthy()
  return (state?.set.find() ?? []).map((d) => ({
    text: editor.state.doc.textBetween(d.from, d.to, ' ', ' '),
    cls: String(d.type.attrs?.class ?? ''),
  }))
}

describe('link/anchor diagnostics (UX-1701)', () => {
  it('marks a link whose reference definition is circular/undefined as unresolved', async () => {
    // [b]: [a] makes the parser produce the literal href "[a]" — honest marking
    // instead of a link that silently looks valid
    const editor = await newEditor('See [b] here.\n\n[b]: [a]')
    const found = diagnostics(editor).filter((d) => d.cls.startsWith('md-link-'))
    expect(found).toEqual([{ text: 'b', cls: 'md-link-unresolved-ref' }])
    // view-only: the model round-trips unchanged
    expect(editor.getMarkdown()).toBe('See [b]([a]) here.')
  })

  it('marks a link to a missing anchor as a dead anchor', async () => {
    const editor = await newEditor('[text](#no-such-anchor)')
    const found = diagnostics(editor).filter((d) => d.cls.startsWith('md-link-'))
    expect(found).toEqual([{ text: 'text', cls: 'md-link-dead-anchor' }])
    expect(editor.getMarkdown()).toBe('[text](#no-such-anchor)')
  })

  it('marks an empty link', async () => {
    const editor = await newEditor('[empty]()')
    const found = diagnostics(editor).filter((d) => d.cls.startsWith('md-link-'))
    expect(found).toEqual([{ text: 'empty', cls: 'md-link-empty' }])
  })

  it('leaves valid links and resolved references unmarked', async () => {
    const editor = await newEditor(
      '[ok](https://example.com) and [ref][real]\n\n[real]: https://example.com/x',
    )
    expect(diagnostics(editor)).toEqual([])
  })

  it('resolves GitHub-style heading slugs, including punctuation stripping', async () => {
    const editor = await newEditor('# My Section!\n\n[go](#my-section) but [bad](#nope)')
    const dead = diagnostics(editor).filter((d) => d.cls === 'md-link-dead-anchor')
    expect(dead).toEqual([{ text: 'bad', cls: 'md-link-dead-anchor' }])
  })

  it('accepts raw-HTML anchors kept by the RawHtml node as targets', async () => {
    const editor = await newEditor('[jump](#section-x)\n\n<a id="section-x"></a>')
    expect(diagnostics(editor)).toEqual([])
  })

  it('hints at unsupported footnotes and toasts once per document', async () => {
    setToastEmitter((toast) => toasts.push(toast))
    const editor = await newEditor('Text with [^1] footnote.\n\n[^1]: Note body.')

    const hints = diagnostics(editor).filter((d) => d.cls === 'md-footnote-hint')
    expect(hints.map((h) => h.text)).toEqual(['[^1]', '[^1]'])
    // 'create' fires on a macrotask after mount — the toast follows it
    await vi.waitFor(() => expect(toasts).toHaveLength(1))
    expect(toasts[0]).toEqual({ text: strings.zh.footnoteUnsupported, kind: 'error' })

    // more footnotes later in the session do not re-toast (one per document)
    const pos = editor.state.doc.content.size - 1
    editor.view.dispatch(editor.state.tr.insertText(' [^2]', pos))
    expect(toasts).toHaveLength(1)
  })

  it('keeps a document without diagnostics silent', async () => {
    setToastEmitter((toast) => toasts.push(toast))
    const editor = await newEditor('# Head\n\nPlain [text](https://example.com) only.')
    expect(diagnostics(editor)).toEqual([])
    await new Promise((resolve) => setTimeout(resolve, 0)) // let 'create' fire
    expect(toasts).toEqual([])
  })
})

describe('diagnostics helpers', () => {
  const anchors = new Set(['my-section', 'my section'])

  it('computes GitHub-style heading slugs', () => {
    expect(headingSlug('My Section!')).toBe('my-section')
    expect(headingSlug('  Hello, World!  ')).toBe('hello-world')
    expect(headingSlug('A  B')).toBe('a--b') // GitHub keeps one hyphen per space
    expect(headingSlug('深度学習')).toBe('深度学習') // unicode letters survive
  })

  it('classifies hrefs without touching valid ones', () => {
    expect(linkIssue('', anchors)).toBe('empty')
    expect(linkIssue('[a]', anchors)).toBe('unresolvedRef')
    expect(linkIssue('#missing', anchors)).toBe('deadAnchor')
    expect(linkIssue('#my-section', anchors)).toBeNull()
    // percent-encoded anchors decode before comparison; malformed stays raw
    expect(linkIssue('#my%20section', anchors)).toBeNull()
    expect(linkIssue('#my%2', anchors)).toBe('deadAnchor')
    expect(linkIssue('https://example.com', anchors)).toBeNull()
    expect(linkIssue('relative/path.md', anchors)).toBeNull()
  })
})

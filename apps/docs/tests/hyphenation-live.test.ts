import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { ParsedDocFull } from '@airy-office/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyHyphenationLive } from '../src/renderer/file-actions'
import { SHORTCUTS } from '../src/renderer/shortcuts'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function emptyEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] },
  })
}

function parsedWith(autoHyphenation?: boolean): ParsedDocFull {
  return {
    styles: new Map(),
    docDefaults: {},
    blocks: [],
    ...(autoHyphenation ? { autoHyphenation: true } : {}),
  } as unknown as ParsedDocFull
}

describe('Layout → Hyphenation live toggle', () => {
  it('sets the lang attribute Chromium needs when on, removes it when off', () => {
    const parsed = parsedWith(true)
    ;(parsed as { docDefaults?: { lang?: string } }).docDefaults = { lang: 'en-US' }
    const editor = emptyEditor()
    applyHyphenationLive(editor, parsed, true)
    expect(editor.view.dom.getAttribute('lang')).toBe('en-US')
    applyHyphenationLive(editor, parsed, false)
    expect(editor.view.dom.getAttribute('lang')).toBeNull()
    editor.destroy()
  })

  it('the regenerated doc CSS carries hyphens:auto exactly for the on state', () => {
    const on = docStyleCss(parsedWith(true))
    const off = docStyleCss(parsedWith(undefined))
    expect(on).toContain('hyphens:auto')
    expect(off).not.toContain('hyphens:auto')
  })

  it('soft hyphen inserts as text and survives the run model', () => {
    const editor = emptyEditor()
    editor.commands.insertContent('hy\u00adphen')
    expect(editor.state.doc.textContent).toBe('hy\u00adphen')
    editor.destroy()
  })

  it('the optional-hyphen shortcut is registered (⌥⌘- / Ctrl+Alt+-)', () => {
    const def = SHORTCUTS.find((s) => s.id === 'soft-hyphen')
    expect(def?.keys).toBe('⌥⌘-')
    expect(def?.win).toBe('Ctrl+Alt+-')
    expect(def?.labelKey).toBe('layoutSoftHyphen')
  })
})

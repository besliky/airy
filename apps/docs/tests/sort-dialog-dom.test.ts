import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { SortDialog } from '../src/renderer/components/SortDialog'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: dictionaries load per locale; the en dictionary this suite
  // asserts against must be loaded first, mirroring the bootstrap-time load.
  return loadLocale('en')
})

/**
 * SortDialog adopts the shared useModalDialog hook; this is the smoke test
 * that the dialog semantics survive adoption: role/aria-modal/labelledby,
 * Escape from a non-input control, the Tab trap across Dropdown triggers,
 * and focus landing back on the trigger after close.
 */

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

/** two paragraphs selected across their full extent: the paragraphs sort scope */
function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'beta' }],
        },
        {
          type: 'docParagraph',
          attrs: { docxIndex: 1 },
          content: [{ type: 'text', text: 'alpha' }],
        },
      ],
    },
  })
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, 1, editor.state.doc.content.size - 1),
    ),
  )
  return editor
}

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function key(element: Element, keyName: string): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key: keyName, bubbles: true, cancelable: true })
  act(() => {
    element.dispatchEvent(event)
  })
  return event
}

describe('SortDialog modal semantics', () => {
  it('renders as a dialog named by its heading', () => {
    const editor = createEditor()
    const { container, unmount } = render(createElement(SortDialog, { editor, onClose: vi.fn() }))
    const modal = container.querySelector('.modal') as HTMLElement
    const title = container.querySelector('.modal h2') as HTMLElement
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe('Sort')
    unmount()
    editor.destroy()
  })

  it('closes on Escape pressed on a button, not just inside the fields', () => {
    const editor = createEditor()
    const onClose = vi.fn()
    const { container, unmount } = render(createElement(SortDialog, { editor, onClose }))
    const cancel = container.querySelector('.modal-actions .btn-ghost') as HTMLElement
    cancel.focus()
    expect(key(cancel, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
    editor.destroy()
  })

  it('wraps Tab from the last control back to the first Dropdown trigger', () => {
    const editor = createEditor()
    const { container, unmount } = render(createElement(SortDialog, { editor, onClose: vi.fn() }))
    const ok = container.querySelector('.modal-actions .btn-primary') as HTMLElement
    ok.focus()
    expect(key(ok, 'Tab').defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(container.querySelector('.gs-dd-btn'))
    unmount()
    editor.destroy()
  })

  it('returns focus to the trigger that opened the dialog', () => {
    const editor = createEditor()
    const trigger = document.createElement('button')
    document.body.appendChild(trigger)
    trigger.focus()
    const { unmount } = render(createElement(SortDialog, { editor, onClose: () => undefined }))
    expect(document.activeElement).not.toBe(trigger)
    unmount()
    expect(document.activeElement).toBe(trigger)
    editor.destroy()
  })
})

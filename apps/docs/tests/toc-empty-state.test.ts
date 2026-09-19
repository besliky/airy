import { beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { TocOptionsModal, TofModal } from '../src/renderer/components/ribbon-references-tab'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: the en dictionary these assertions read must load first
  return loadLocale('en')
})

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

/**
 * UX-1010: the TOC options / Table of Figures dialogs report their empty
 * states inline (a .modal-error role="alert" line inside the open dialog)
 * instead of a blocking window.alert on top of the modal. The dialog stays
 * open so the user can cancel knowingly.
 */

function emptyEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
  })
}

function renderModal(element: React.ReactElement): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return { container, root }
}

function clickInsert(container: HTMLElement): void {
  const btn = [...container.querySelectorAll('button')].find((b) =>
    ['Insert', 'Update'].includes(b.textContent ?? ''),
  )
  expect(btn).toBeTruthy()
  act(() => btn!.dispatchEvent(new MouseEvent('click', { bubbles: true })))
}

describe('TOC/ToF dialogs empty states stay inline (UX-1010)', () => {
  it('TofModal shows a .modal-error line instead of window.alert', () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => undefined)
    const editor = emptyEditor()
    const { container, root } = renderModal(
      createElement(TofModal, { editor, blocks: [], onClose: vi.fn() }),
    )
    clickInsert(container)
    const error = container.querySelector('.modal-error')
    expect(error).not.toBeNull()
    expect(error!.getAttribute('role')).toBe('alert')
    expect(alert).not.toHaveBeenCalled()
    // the dialog stays open — the message is a state, not a dismissal
    expect(container.querySelector('.modal')).not.toBeNull()
    act(() => root.unmount())
    editor.destroy()
    alert.mockRestore()
  })

  it('TocOptionsModal shows a .modal-error line instead of window.alert', () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => undefined)
    const editor = emptyEditor()
    const { container, root } = renderModal(
      createElement(TocOptionsModal, { editor, headingPages: undefined, onClose: vi.fn() }),
    )
    clickInsert(container)
    const error = container.querySelector('.modal-error')
    expect(error).not.toBeNull()
    expect(error!.getAttribute('role')).toBe('alert')
    expect(alert).not.toHaveBeenCalled()
    act(() => root.unmount())
    editor.destroy()
    alert.mockRestore()
  })
})

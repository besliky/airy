import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, setModuleLang, t } from '../src/renderer/i18n/locale'
import { PdfExportDialog } from '../src/renderer/components/PdfExportDialog'

/**
 * UX-1101 migration smoke: the slides PDF-export layout chooser moved from
 * the local useModalKeys shim onto the shared useModalDialog hook — dialog
 * role / aria-modal / aria-labelledby naming, initial focus, focus trap,
 * stopped Escape and focus return. The class-level "no unhooked modal" gate
 * lives in modal-dialog-contract.test.ts.
 */
beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { slidesApi: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

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

describe('PdfExportDialog modal semantics (UX-1101)', () => {
  const renderDialog = (onClose = vi.fn()) =>
    render(createElement(PdfExportDialog, { slideCount: 3, onExport: vi.fn(), onClose }))

  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = renderDialog()
    const modal = container.querySelector('.modal') as HTMLElement
    const title = container.querySelector('.modal h2') as HTMLElement
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe(t('ribbonFileExportPdf'))
    unmount()
  })

  it('focuses the first layout radio on open and closes on Escape from it', () => {
    const onClose = vi.fn()
    const { container, unmount } = renderDialog(onClose)
    const firstRadio = container.querySelector<HTMLInputElement>('input[type="radio"]')
    expect(document.activeElement).toBe(firstRadio)
    expect(key(firstRadio as Element, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('wraps Tab at the dialog ends (focus trap)', () => {
    const { container, unmount } = renderDialog()
    const firstRadio = container.querySelector('input[type="radio"]') as HTMLElement
    const lastButton = [...container.querySelectorAll('button')].pop() as HTMLElement
    lastButton.focus()
    expect(key(lastButton, 'Tab').defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(firstRadio)
    unmount()
  })

  it('returns focus to the trigger that opened the dialog', () => {
    const trigger = document.createElement('button')
    document.body.appendChild(trigger)
    trigger.focus()
    const { unmount } = renderDialog()
    expect(document.activeElement).not.toBe(trigger)
    unmount()
    expect(document.activeElement).toBe(trigger)
    trigger.remove()
  })
})

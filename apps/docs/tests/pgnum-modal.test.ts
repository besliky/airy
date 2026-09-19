import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'
import { PgNumFormatModal } from '../src/renderer/App'

/**
 * UX-1207 migration smoke: the Page Number Format modal sat in App.tsx's
 * root, where the docs modal contract (which scanned only components/)
 * could not see it — a legacy dialog without Escape, focus trap or dialog
 * semantics. It now lives in a local component on the shared useModalDialog
 * hook; the class-level gate lives in modal-dialog-contract.test.ts (which
 * sweeps the whole renderer since this change).
 */
beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  return loadLocale('en')
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

function renderModal(onClose = vi.fn()): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(PgNumFormatModal, {
          fmt: 'decimal',
          start: '',
          sectionHint: t('appPgNumHintBlank'),
          onFmt: vi.fn(),
          onStart: vi.fn(),
          onApply: vi.fn(),
          onClose,
        }),
      }),
    ),
  )
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

describe('PgNumFormatModal semantics (UX-1207)', () => {
  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = renderModal()
    const modal = container.querySelector('.modal') as HTMLElement
    const title = container.querySelector('.modal h2') as HTMLElement
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe(t('appPageNumFormatTitle'))
    unmount()
  })

  it('focuses the first control on open and closes on Escape from it', () => {
    const onClose = vi.fn()
    const { container, unmount } = renderModal(onClose)
    // the Dropdown's trigger button is the first focusable control
    const first = container.querySelector<HTMLElement>('.gs-dd-btn')!
    expect(document.activeElement).toBe(first)
    expect(key(first, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('returns focus to the trigger that opened the dialog', () => {
    const trigger = document.createElement('button')
    document.body.appendChild(trigger)
    trigger.focus()
    const { unmount } = renderModal()
    expect(document.activeElement).not.toBe(trigger)
    unmount()
    expect(document.activeElement).toBe(trigger)
    trigger.remove()
  })
})

import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { LineNumbersDialog } from '../src/renderer/components/LineNumbersDialog'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: dictionaries load per locale; the en dictionary this suite
  // asserts against must be loaded first, mirroring the bootstrap-time load.
  return loadLocale('en')
})

/**
 * LineNumbersDialog adopts the shared useModalDialog hook (UX-904); this is
 * the smoke test that the dialog semantics survive adoption: role/aria-modal/
 * labelledby, Escape from a plain button, the Tab trap across the radio row,
 * and focus landing back on the trigger after close.
 */

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

function render(props: Partial<React.ComponentProps<typeof LineNumbersDialog>> = {}): {
  container: HTMLElement
  unmount: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(LineNumbersDialog, {
          value: undefined,
          onApply: vi.fn(),
          onClose: vi.fn(),
          ...props,
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

describe('LineNumbersDialog modal semantics (UX-904)', () => {
  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = render()
    const modal = container.querySelector('.modal') as HTMLElement
    const title = container.querySelector('.modal h2') as HTMLElement
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe('Line Numbers')
    unmount()
  })

  it('closes on Escape pressed on a button, not just inside the fields', () => {
    const onClose = vi.fn()
    const { container, unmount } = render({ onClose })
    const cancel = container.querySelector('.modal-actions .btn-ghost') as HTMLElement
    cancel.focus()
    expect(key(cancel, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('wraps Tab from the last control back to the first radio', () => {
    const { container, unmount } = render()
    const ok = container.querySelector('.modal-actions .btn-primary') as HTMLElement
    ok.focus()
    expect(key(ok, 'Tab').defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(container.querySelector('.ln-radio input'))
    unmount()
  })

  it('returns focus to the trigger that opened the dialog', () => {
    const trigger = document.createElement('button')
    document.body.appendChild(trigger)
    trigger.focus()
    const { unmount } = render({ onClose: () => undefined })
    expect(document.activeElement).not.toBe(trigger)
    unmount()
    expect(document.activeElement).toBe(trigger)
  })
})

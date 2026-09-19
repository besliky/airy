// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'
import { TextToColumnsDialog } from '../src/renderer/TextToColumnsDialog'
import { OutlineSettingsDialog } from '../src/renderer/OutlineSettingsDialog'

/**
 * UX-1101 migration smoke: the PR #67 sheets dialogs (Text to Columns and
 * Outline Settings) shipped with a bare role="dialog" and no keyboard
 * handling at all; both moved onto the shared useModalDialog hook — dialog
 * role / aria-modal / aria-labelledby naming, initial focus, stopped Escape
 * and focus return. The class-level "no unhooked dialog" gate lives in
 * modal-dialog-contract.test.ts.
 */
beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-901: dictionaries load per locale; the en dictionary this suite
  // asserts against must be loaded first, mirroring the bootstrap-time load.
  return loadLocale('en')
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktopApi: { onLanguageChanged: () => () => undefined } })

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

/** role="dialog" named by its visible header */
const expectDialogSemantics = (
  container: HTMLElement,
  box: string,
  titleKey: Parameters<typeof t>[0],
) => {
  const dialog = container.querySelector(box) as HTMLElement
  const header = container.querySelector(`${box} > header`) as HTMLElement
  expect(dialog.getAttribute('role')).toBe('dialog')
  expect(dialog.getAttribute('aria-modal')).toBe('true')
  expect(dialog.getAttribute('aria-labelledby')).toBe(header.id)
  expect(header.textContent).toBe(t(titleKey))
}

describe('TextToColumnsDialog modal semantics (UX-1101)', () => {
  const renderDialog = (onClose = vi.fn()) =>
    render(
      createElement(TextToColumnsDialog, {
        source: { rows: ['a,b', 'c,d'], destinationLabel: 'B1' },
        onApply: vi.fn().mockReturnValue(null),
        onClose,
      }),
    )

  it('renders as a dialog named by its header', () => {
    const { container, unmount } = renderDialog()
    expectDialogSemantics(container, '.format-cells-dialog', 'dlgT2cTitle')
    unmount()
  })

  it('focuses the first mode radio on open and closes on Escape from it', () => {
    const onClose = vi.fn()
    const { container, unmount } = renderDialog(onClose)
    const firstRadio = container.querySelector<HTMLInputElement>('input[type="radio"]')
    expect(document.activeElement).toBe(firstRadio)
    expect(key(firstRadio as Element, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
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

describe('OutlineSettingsDialog modal semantics (UX-1101)', () => {
  const renderDialog = (onClose = vi.fn()) =>
    render(
      createElement(OutlineSettingsDialog, {
        initial: { summaryBelow: true, summaryRight: true },
        onApply: vi.fn().mockReturnValue(null),
        onClose,
      }),
    )

  it('renders as a dialog named by its header', () => {
    const { container, unmount } = renderDialog()
    expectDialogSemantics(container, '.format-cells-dialog', 'dlgOutlineSettingsTitle')
    unmount()
  })

  it('closes on Escape from the first radio', () => {
    const onClose = vi.fn()
    const { container, unmount } = renderDialog(onClose)
    const firstRadio = container.querySelector('input[type="radio"]') as HTMLElement
    expect(document.activeElement).toBe(firstRadio)
    expect(key(firstRadio, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })
})

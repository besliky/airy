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

describe('Text to Columns: Enter applies, overwriting asks first (UX-1108)', () => {
  const renderDialog = (overwrites: boolean) => {
    const onApply = vi.fn().mockReturnValue(null)
    const onClose = vi.fn()
    const rendered = render(
      createElement(TextToColumnsDialog, {
        source: { rows: ['a,b', 'c,d'], destinationLabel: 'A1' },
        onApply,
        onDestinationOverwrites: () => overwrites,
        onClose,
      }),
    )
    return { onApply, onClose, ...rendered }
  }

  const destinationInput = (container: HTMLElement): HTMLInputElement =>
    container.querySelector<HTMLInputElement>('.t2c-destination-input') as HTMLInputElement

  const typeInto = (input: HTMLInputElement, value: string): void => {
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
      setter?.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
  }

  it('Enter in the destination field applies on the first press when nothing is overwritten', () => {
    const { onApply, onClose, container, unmount } = renderDialog(false)
    const destination = destinationInput(container)
    expect(key(destination, 'Enter').defaultPrevented).toBe(true)
    expect(onApply).toHaveBeenCalledTimes(1)
    expect(onApply.mock.calls[0]?.[0]).toMatchObject({ destination: null })
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('Enter arms the overwrite confirm; the second press applies', () => {
    const { onApply, onClose, container, unmount } = renderDialog(true)
    const destination = destinationInput(container)
    key(destination, 'Enter')
    expect(onApply).not.toHaveBeenCalled()
    expect(container.textContent).toContain(t('dlgT2cOverwriteNote'))
    const ok = [...container.querySelectorAll('button')].find((button) =>
      button.className.includes('primary-action'),
    ) as HTMLElement
    expect(ok.textContent).toBe(t('dlgT2cOverwriteConfirm'))
    key(destination, 'Enter')
    expect(onApply).toHaveBeenCalledTimes(1)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('editing the destination disarms the confirm', () => {
    const { onApply, container, unmount } = renderDialog(true)
    const destination = destinationInput(container)
    key(destination, 'Enter')
    expect(container.textContent).toContain(t('dlgT2cOverwriteNote'))
    typeInto(destination, 'E1')
    expect(container.textContent).not.toContain(t('dlgT2cOverwriteNote'))
    key(destination, 'Enter')
    // the probe runs again for the edited destination — still overwriting
    // here, so the wizard asks again instead of applying
    expect(onApply).not.toHaveBeenCalled()
    expect(container.textContent).toContain(t('dlgT2cOverwriteNote'))
    unmount()
  })

  it('Enter does not apply while the config is invalid', () => {
    const { onApply, container, unmount } = renderDialog(false)
    // switch to fixed width so the breaks field exists, then make it invalid
    const radios = [...container.querySelectorAll<HTMLInputElement>('input[name="t2c-mode"]')]
    act(() => radios[1]?.click())
    const breaks = container.querySelector<HTMLInputElement>(
      '.t2c-breaks-input',
    ) as HTMLInputElement
    typeInto(breaks, 'x')
    key(breaks, 'Enter')
    expect(onApply).not.toHaveBeenCalled()
    unmount()
  })

  it('a promise probe (streamed file floor) arms the confirm, then applies', async () => {
    // BUG-1312: the probe reads the destination through the sidecar on
    // streamed workbooks, so its answer arrives asynchronously
    const onApply = vi.fn().mockReturnValue(null)
    const onClose = vi.fn()
    let resolveProbe: (asks: boolean) => void = () => {}
    const onDestinationOverwrites = () =>
      new Promise<boolean>((resolve) => {
        resolveProbe = resolve
      })
    const { container, unmount } = render(
      createElement(TextToColumnsDialog, {
        source: { rows: ['a,b', 'c,d'], destinationLabel: 'A1' },
        onApply,
        onDestinationOverwrites,
        onClose,
      }),
    )
    const destination = destinationInput(container)
    key(destination, 'Enter')
    // while the probe is in flight the commit button is blocked
    const ok = [...container.querySelectorAll('button')].find((button) =>
      button.className.includes('primary-action'),
    ) as HTMLButtonElement
    expect(ok.disabled).toBe(true)
    expect(onApply).not.toHaveBeenCalled()
    await act(async () => {
      resolveProbe(true)
    })
    expect(onApply).not.toHaveBeenCalled()
    expect(container.textContent).toContain(t('dlgT2cOverwriteNote'))
    expect(ok.disabled).toBe(false)
    // the armed confirm skips the probe: the second commit applies directly
    act(() => ok.click())
    expect(onApply).toHaveBeenCalledTimes(1)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('a failed apply disarms the confirm so the retry re-runs the probe', () => {
    const onApply = vi.fn().mockReturnValueOnce('boom').mockReturnValue(null)
    const onDestinationOverwrites = vi.fn(() => true)
    const { container, unmount } = render(
      createElement(TextToColumnsDialog, {
        source: { rows: ['a,b', 'c,d'], destinationLabel: 'A1' },
        onApply,
        onDestinationOverwrites,
        onClose: () => {},
      }),
    )
    const destination = destinationInput(container)
    key(destination, 'Enter')
    expect(container.textContent).toContain(t('dlgT2cOverwriteNote'))
    key(destination, 'Enter')
    expect(onApply).toHaveBeenCalledTimes(1)
    expect(container.textContent).toContain('boom')
    // the confirm is gone: the next Enter probes again instead of applying
    expect(container.textContent).not.toContain(t('dlgT2cOverwriteNote'))
    key(destination, 'Enter')
    expect(onApply).toHaveBeenCalledTimes(1)
    // probe #1 armed the confirm, the armed commit skipped it, the retry
    // after the failure ran it again
    expect(onDestinationOverwrites).toHaveBeenCalledTimes(2)
    unmount()
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

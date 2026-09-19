/**
 * AltTextDialog UI contract (TEST-1001, PAR-115): the shared alt-text editor
 * for pictures, shapes and tables. The engine side (wp:docPr title/descr and
 * w:tblCaption/w:tblDescription round-trip, XML escaping) is covered by
 * docx-engine's picture-effects suite; these tests pin the dialog layer that
 * had no direct coverage: modal semantics via the shared useModalDialog hook,
 * both fields editing and passing their value through onApply, and the
 * apply affordances (primary button, Enter in the title field, Cancel).
 */
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { AltTextDialog, type AltTextValue } from '../src/renderer/components/AltTextDialog'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: dictionaries load per locale; the en dictionary this suite
  // asserts against must be loaded first, mirroring the bootstrap-time load.
  Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
  setModuleLang('en')
  return loadLocale('en')
})

// The dialog's focus/Escape effects live on the window; every test must
// unmount its dialog even when an assertion fails, or a stale listener eats
// the next test's keys.
const mounted: Array<() => void> = []
beforeEach(() => mounted.splice(0).forEach((unmount) => unmount()))

function mountDialog(initial: AltTextValue = { title: '', description: '' }): {
  host: HTMLElement
  onApply: ReturnType<typeof vi.fn>
  onCancel: ReturnType<typeof vi.fn>
  titleInput: () => HTMLInputElement
  descriptionInput: () => HTMLTextAreaElement
  applyButton: () => HTMLButtonElement
  cancel: () => void
} {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const onApply = vi.fn<(value: AltTextValue) => void>()
  const onCancel = vi.fn()
  let root!: Root
  act(() => {
    root = createRoot(host)
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(AltTextDialog, { initial, onApply, onCancel }),
      }),
    )
  })
  mounted.push(() => {
    act(() => root.unmount())
    host.remove()
  })
  const titleInput = () => host.querySelector<HTMLInputElement>('.alt-text-field input')!
  const descriptionInput = () => host.querySelector<HTMLTextAreaElement>('textarea')!
  const applyButton = () =>
    [...host.querySelectorAll<HTMLButtonElement>('.modal-actions button')].find((b) =>
      b.className.includes('primary'),
    )!
  return {
    host,
    onApply,
    onCancel,
    titleInput,
    descriptionInput,
    applyButton,
    cancel: () =>
      act(() => {
        host.querySelectorAll<HTMLButtonElement>('.modal-actions button')[0]!.click()
      }),
  }
}

function setValue(input: HTMLInputElement | HTMLTextAreaElement, value: string): void {
  const proto = input instanceof HTMLInputElement ? HTMLInputElement : HTMLTextAreaElement
  const setter = Object.getOwnPropertyDescriptor(proto.prototype, 'value')!.set!
  act(() => {
    setter.call(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function key(element: Element, keyName: string): KeyboardEvent {
  const event = new KeyboardEvent('keydown', { key: keyName, bubbles: true, cancelable: true })
  act(() => {
    element.dispatchEvent(event)
  })
  return event
}

describe('AltTextDialog modal semantics', () => {
  it('renders as a dialog named by its heading and focuses the first field', () => {
    const d = mountDialog()
    const modal = d.host.querySelector<HTMLElement>('.modal')!
    const title = d.host.querySelector<HTMLElement>('.modal h2')!
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe('Alt Text')
    expect(d.titleInput().maxLength).toBe(255)
    // focus lands inside the dialog on open (the title input is autoFocus)
    expect(d.host.contains(document.activeElement)).toBe(true)
  })

  it('closes on Escape from a field and wraps Tab at the dialog edges', () => {
    const d = mountDialog()
    expect(key(d.titleInput(), 'Escape').defaultPrevented).toBe(true)
    expect(d.onCancel).toHaveBeenCalledTimes(1)

    const apply = d.applyButton()
    apply.focus()
    expect(key(apply, 'Tab').defaultPrevented).toBe(true)
    // wrapped back inside the dialog instead of escaping behind the backdrop
    expect(d.host.contains(document.activeElement)).toBe(true)
  })
})

describe('AltTextDialog value flow', () => {
  it('passes the edited title and description through onApply untouched', () => {
    // XML-special and quote characters must arrive raw: escaping is the
    // engine writer's job (docPr / tblDescription), never the dialog's
    const d = mountDialog({ title: 'Chart', description: 'Sales by region' })
    setValue(d.titleInput(), '<FY>24 & "Q3"')
    setValue(d.descriptionInput(), 'Revenue \u00e9\u00e0 chart with <amp> & tags')
    act(() => d.applyButton().click())
    expect(d.onApply).toHaveBeenCalledTimes(1)
    expect(d.onApply).toHaveBeenCalledWith({
      title: '<FY>24 & "Q3"',
      description: 'Revenue \u00e9\u00e0 chart with <amp> & tags',
    })
    expect(d.onCancel).not.toHaveBeenCalled()
  })

  it('applies on Enter in the title field; Cancel leaves the value unreported', () => {
    const d = mountDialog({ title: '', description: 'kept' })
    setValue(d.titleInput(), 'entered')
    key(d.titleInput(), 'Enter')
    expect(d.onApply).toHaveBeenCalledWith({ title: 'entered', description: 'kept' })

    const e = mountDialog({ title: 'stays', description: '' })
    e.cancel()
    expect(e.onCancel).toHaveBeenCalledTimes(1)
    expect(e.onApply).not.toHaveBeenCalled()
  })
})

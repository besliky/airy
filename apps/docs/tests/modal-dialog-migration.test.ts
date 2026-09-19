import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { DEFAULT_SECTION } from '@airy-office/docx-engine'
import { LocaleProvider, loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'
import { AltTextDialog } from '../src/renderer/components/AltTextDialog'
import { ColumnsDialog } from '../src/renderer/components/ColumnsDialog'
import { NoteOptionsDialog } from '../src/renderer/components/NoteOptionsDialog'
import { CompressPicturesDialog, CutoutDialog } from '../src/renderer/components/PictureDialogs'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: dictionaries load per locale; the en dictionary this suite
  // asserts against must be loaded first, mirroring the bootstrap-time load.
  return loadLocale('en')
})

/**
 * UX-1001 migration smoke tests: the nine v0.12.0 dialogs that shipped
 * without the shared useModalDialog hook moved onto it. These are the
 * per-shape representatives (plain form dialog, legacy useModalKeys dialog,
 * references-tab dialog, picture dialogs with an image-loading lifecycle) —
 * the rest share the identical three-spread adoption, and the class-level
 * "no unhooked modal" gate lives in modal-dialog-contract.test.ts.
 */

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

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

/** role="dialog" named by its visible heading, focus-trapped */
const expectDialogSemantics = (container: HTMLElement, titleKey: Parameters<typeof t>[0]) => {
  const modal = container.querySelector('.modal') as HTMLElement
  const title = container.querySelector('.modal h2') as HTMLElement
  expect(modal.getAttribute('role')).toBe('dialog')
  expect(modal.getAttribute('aria-modal')).toBe('true')
  expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
  expect(title.textContent).toBe(t(titleKey))
}

describe('AltTextDialog modal semantics (UX-1001)', () => {
  const renderDialog = (onCancel = vi.fn()) =>
    render(
      createElement(AltTextDialog, {
        initial: { title: '', description: '' },
        onApply: vi.fn(),
        onCancel,
      }),
    )

  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = renderDialog()
    expectDialogSemantics(container, 'ribbonAltText')
    unmount()
  })

  it('focuses the title field on open and closes on Escape from it', () => {
    const onCancel = vi.fn()
    const { container, unmount } = renderDialog(onCancel)
    const titleField = container.querySelector('.alt-text-field input') as HTMLElement
    expect(document.activeElement).toBe(titleField)
    expect(key(titleField, 'Escape').defaultPrevented).toBe(true)
    expect(onCancel).toHaveBeenCalledTimes(1)
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
  })
})

describe('ColumnsDialog modal semantics (UX-1001)', () => {
  const renderDialog = (onClose = vi.fn()) =>
    render(
      createElement(ColumnsDialog, { section: { ...DEFAULT_SECTION }, onApply: vi.fn(), onClose }),
    )

  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = renderDialog()
    expectDialogSemantics(container, 'layoutColsDialogTitle')
    unmount()
  })

  it('closes on Escape pressed on a preset button, not just inside the fields', () => {
    const onClose = vi.fn()
    const { container, unmount } = renderDialog(onClose)
    const preset = container.querySelector('.cols-preset') as HTMLElement
    expect(key(preset, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('wraps Tab from OK back to the first preset button (focus trap)', () => {
    const { container, unmount } = renderDialog()
    const ok = container.querySelector('.modal-actions .btn-primary') as HTMLElement
    ok.focus()
    expect(key(ok, 'Tab').defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(container.querySelector('.cols-preset'))
    unmount()
  })
})

describe('NoteOptionsDialog modal semantics (UX-1001)', () => {
  const renderDialog = (onClose = vi.fn()) =>
    render(
      createElement(NoteOptionsDialog, {
        value: {},
        hasSelection: false,
        selectionKind: undefined,
        onApply: vi.fn(),
        onConvert: vi.fn(),
        onClose,
      }),
    )

  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = renderDialog()
    expectDialogSemantics(container, 'refsNoteOptionsTitle')
    unmount()
  })

  it('focus lands inside on open and Escape closes from the convert row', () => {
    const onClose = vi.fn()
    const { container, unmount } = renderDialog(onClose)
    const backdrop = container.querySelector('.modal-backdrop') as HTMLElement
    expect(backdrop.contains(document.activeElement)).toBe(true)
    const convert = container.querySelector('.modal-actions .btn-ghost') as HTMLElement
    expect(key(convert, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })
})

describe('picture dialogs modal semantics (UX-1001)', () => {
  const PIXEL =
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

  it('CompressPicturesDialog renders as a named dialog and closes on Escape', () => {
    const onCancel = vi.fn()
    const { container, unmount } = render(
      createElement(CompressPicturesDialog, {
        dataUrl: PIXEL,
        displayWidthPx: 400,
        displayHeightPx: 300,
        crop: null,
        onApply: vi.fn(),
        onCancel,
      }),
    )
    expectDialogSemantics(container, 'ribbonCompressPictures')
    const cancel = [...container.querySelectorAll('button')].at(-1) as HTMLElement
    expect(key(cancel, 'Escape').defaultPrevented).toBe(true)
    expect(onCancel).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('CutoutDialog renders as a named dialog and its Escape stays inside (no window leak)', () => {
    const onCancel = vi.fn()
    const { container, unmount } = render(
      createElement(CutoutDialog, { dataUrl: PIXEL, onApply: vi.fn(), onCancel }),
    )
    expectDialogSemantics(container, 'ribbonRemoveBg')
    // the dialog opens with the image still decoding: the enabled Cancel
    // button is the first focusable, Escape there must close the dialog
    const cancel = container.querySelector('.modal-actions button') as HTMLElement
    expect(document.activeElement).toBe(cancel)
    const windowEsc = vi.fn()
    window.addEventListener('keydown', windowEsc)
    expect(key(cancel, 'Escape').defaultPrevented).toBe(true)
    expect(onCancel).toHaveBeenCalledTimes(1)
    // the key was stopped before app-global Escape listeners (read mode)
    // could see it — the v0.12.0 regression this migration fixes
    expect(windowEsc).not.toHaveBeenCalled()
    window.removeEventListener('keydown', windowEsc)
    unmount()
  })
})

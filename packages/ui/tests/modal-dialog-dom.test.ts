// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement, useState } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { useModalDialog } from '../src/modal-dialog'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

/* Mounts the smallest dialog the apps build: a backdrop with the backdrop
 * props, a dialog box with the dialog props, a heading with the title props,
 * and a few focusable controls (plus a disabled one that the trap must skip).
 * Closing unmounts the dialog, like every adopter's conditional render. */
function Probe({ onClose }: { onClose: () => void }) {
  const [open, setOpen] = useState(true)
  const dialog = useModalDialog(() => {
    onClose()
    setOpen(false)
  })
  if (!open) return null
  return createElement(
    'div',
    { className: 'backdrop', ...dialog.backdropProps },
    createElement(
      'div',
      { className: 'modal', ...dialog.dialogProps },
      createElement('h2', dialog.titleProps, 'Dialog title'),
      createElement('input', { className: 'first' }),
      createElement('button', { disabled: true }, 'skipped'),
      createElement('button', { className: 'mid' }, 'Do'),
      createElement('a', { href: '#last', className: 'last' }, 'last'),
    ),
  )
}

function mount(onClose: () => void): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(Probe, { onClose })))
  return { root, container }
}

function press(element: Element, keyName: string, shiftKey = false): KeyboardEvent {
  const event = new KeyboardEvent('keydown', {
    key: keyName,
    shiftKey,
    bubbles: true,
    cancelable: true,
  })
  act(() => {
    element.dispatchEvent(event)
  })
  return event
}

describe('useModalDialog', () => {
  it('exposes the dialog ARIA structure named by its heading', () => {
    const { root, container } = mount(vi.fn())
    const modal = container.querySelector('.modal') as HTMLElement
    const title = container.querySelector('h2') as HTMLElement
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe('Dialog title')
    act(() => root.unmount())
  })

  it('focuses the first focusable control on mount', () => {
    const { root, container } = mount(vi.fn())
    expect(document.activeElement).toBe(container.querySelector('.first'))
    act(() => root.unmount())
  })

  it('wraps Tab and Shift+Tab at the dialog ends and leaves middle tabs alone', () => {
    const { root, container } = mount(vi.fn())
    const first = container.querySelector('.first') as HTMLElement
    const mid = container.querySelector('.mid') as HTMLElement
    const last = container.querySelector('.last') as HTMLElement

    last.focus()
    expect(press(last, 'Tab').defaultPrevented).toBe(true)
    expect(document.activeElement).toBe(first)

    first.focus()
    press(first, 'Tab', true)
    expect(document.activeElement).toBe(last)

    // a control in the middle keeps the browser's own Tab handling
    mid.focus()
    expect(press(mid, 'Tab').defaultPrevented).toBe(false)
    act(() => root.unmount())
  })

  it('closes on Escape from an inner control and keeps the key from app-global listeners', () => {
    const onClose = vi.fn()
    const { root, container } = mount(onClose)
    const globalEscape = vi.fn()
    window.addEventListener('keydown', globalEscape)
    const mid = container.querySelector('.mid') as HTMLElement
    mid.focus()
    expect(press(mid, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(globalEscape).not.toHaveBeenCalled()
    // the adopter unmounted the dialog on close
    expect(container.querySelector('.modal')).toBeNull()
    window.removeEventListener('keydown', globalEscape)
    act(() => root.unmount())
  })

  it('closes on Escape while focus sits outside the dialog', () => {
    const onClose = vi.fn()
    const { root, container } = mount(onClose)
    ;(document.activeElement as HTMLElement | null)?.blur()
    expect(document.activeElement).toBe(document.body)
    expect(press(document.body, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(container.querySelector('.modal')).toBeNull()
    act(() => root.unmount())
  })

  it('stands down on a press another layer claimed — one Esc, one layer (BUG-1320)', () => {
    // The ribbon popover's window-capture handler is registered BEFORE the
    // modal mounts and claims Escape with preventDefault+stopPropagation —
    // which does not stop same-node listeners. The fallback used to close the
    // dialog with the same press: one Esc dismissed two layers whenever focus
    // sat outside the modal box (body), where the defer predicate missed.
    const onClose = vi.fn()
    const popoverClosed = vi.fn()
    const popoverClaim = (e: KeyboardEvent) => {
      if (e.key !== 'Escape' || e.defaultPrevented) return
      e.preventDefault()
      e.stopPropagation()
      popoverClosed()
    }
    window.addEventListener('keydown', popoverClaim, true)
    try {
      const { root, container } = mount(onClose)
      ;(document.activeElement as HTMLElement | null)?.blur()
      expect(document.activeElement).toBe(document.body)
      const event = press(document.body, 'Escape')
      // the popover took the press; the modal stays open for the next one
      expect(popoverClosed).toHaveBeenCalledTimes(1)
      expect(event.defaultPrevented).toBe(true)
      expect(onClose).not.toHaveBeenCalled()
      expect(container.querySelector('.modal')).not.toBeNull()
      act(() => root.unmount())
    } finally {
      window.removeEventListener('keydown', popoverClaim, true)
    }
    // the popover is gone now: the next press closes the dialog alone
    const { root, container } = mount(onClose)
    ;(document.activeElement as HTMLElement | null)?.blur()
    expect(press(document.body, 'Escape').defaultPrevented).toBe(true)
    expect(onClose).toHaveBeenCalledTimes(1)
    expect(container.querySelector('.modal')).toBeNull()
    act(() => root.unmount())
  })

  it('returns focus to the trigger when the dialog unmounts', () => {
    const trigger = document.createElement('button')
    document.body.appendChild(trigger)
    trigger.focus()
    const { root } = mount(vi.fn())
    // initial focus moved into the dialog…
    expect(document.activeElement).not.toBe(trigger)
    act(() => root.unmount())
    // …and unmounting hands it back to the trigger
    expect(document.activeElement).toBe(trigger)
  })
})

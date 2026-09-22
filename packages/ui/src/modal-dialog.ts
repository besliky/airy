/**
 * WAI-ARIA modal dialog pattern for the suite's overlay dialogs: the box is a
 * `role="dialog" aria-modal` named by its visible heading, focus is trapped
 * inside (Tab/Shift+Tab wrap at the ends), Escape closes from wherever focus
 * sits, and on unmount focus returns to the element that opened the dialog
 * (its trigger button). The first form control gets focus on mount unless
 * something inside already has it (an `autoFocus` field wins).
 *
 * Escape is handled twice on purpose. Key presses inside the dialog bubble to
 * the backdrop's React handler, where inner widgets can still claim the key
 * first (an open Dropdown list closes itself and stops propagation, so it does
 * not take the modal with it). A window-level capture fallback only fires when
 * focus is NOT inside the dialog, so Escape still closes the dialog when focus
 * sits on the page behind it — and is stopped before app-global Escape
 * listeners (read mode, thumbnail navigation) see it. The fallback respects a
 * press another layer already claimed (defaultPrevented, BUG-1320): one
 * Escape dismisses exactly one layer.
 */
import { useCallback, useEffect, useId, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react'

/**
 * Focusable dialog controls in DOM order. Dropdown listbox options carry
 * `tabindex="-1"` on purpose (menu-button pattern: focus lives on the
 * trigger) and are excluded along with disabled controls.
 */
const FOCUSABLE = [
  'a[href]',
  'button:not([disabled])',
  'input:not([disabled])',
  'select:not([disabled])',
  'textarea:not([disabled])',
  '[contenteditable]:not([contenteditable="false"])',
  '[tabindex]:not([tabindex="-1"])',
].join(', ')

export interface ModalDialogController {
  /** Spread on the backdrop wrapper (the element that dims the page). */
  readonly backdropProps: {
    readonly ref: RefObject<HTMLDivElement | null>
    readonly onKeyDown: (e: ReactKeyboardEvent<HTMLDivElement>) => void
  }
  /** Spread on the dialog box element (the `.modal` / `.pdf-modal` div). */
  readonly dialogProps: {
    readonly role: 'dialog'
    readonly 'aria-modal': true
    readonly 'aria-labelledby': string
  }
  /** Spread on the visible heading so the dialog can name itself by it. */
  readonly titleProps: {
    readonly id: string
  }
}

export function useModalDialog(onClose: () => void): ModalDialogController {
  const ref = useRef<HTMLDivElement>(null)
  const titleId = useId()
  // latest-ref: the key handlers stay referentially stable across the host's
  // re-renders (modal dialogs re-render on every keystroke) without going stale
  const closeRef = useRef(onClose)
  useEffect(() => {
    closeRef.current = onClose
  })

  // initial focus in, focus back to the trigger on unmount (the parent
  // unmounts the dialog on close, so every close path restores focus)
  useEffect(() => {
    const root = ref.current
    const trigger = document.activeElement
    if (root && !root.contains(trigger)) {
      const first = root.querySelector<HTMLElement>(FOCUSABLE)
      ;(first ?? root).focus()
    }
    return () => {
      if (trigger instanceof HTMLElement && trigger.isConnected) trigger.focus()
    }
  }, [])

  // Escape with focus outside the dialog (body, or chrome the focus trap
  // missed): capture phase, stopped before app-global listeners see it.
  // A press another layer already claimed (defaultPrevented — e.g. the
  // ribbon popover's window-capture handler, BUG-1320) belongs to that
  // layer: "one Escape dismisses one layer", so this fallback stands down
  // and the next press closes the dialog.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== 'Escape' || e.isComposing) return
      if (e.defaultPrevented) return
      if (ref.current?.contains(document.activeElement)) return
      e.preventDefault()
      e.stopPropagation()
      closeRef.current()
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [])

  const onKeyDown = useCallback((e: ReactKeyboardEvent<HTMLDivElement>) => {
    if (e.key === 'Escape') {
      // mid-IME-composition Escape cancels the composition, not the dialog
      if (e.nativeEvent.isComposing) return
      e.preventDefault()
      // keep the key from reaching global listeners (read mode and friends)
      e.stopPropagation()
      closeRef.current()
      return
    }
    if (e.key !== 'Tab' || e.altKey || e.ctrlKey || e.metaKey) return
    const root = ref.current
    if (!root) return
    const focusables = Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE)).filter(
      // jsdom has no layout: there everything counts as rendered
      (el) => typeof el.checkVisibility !== 'function' || el.checkVisibility(),
    )
    const first = focusables[0]
    const last = focusables[focusables.length - 1]
    if (!first || !last) return
    const active = document.activeElement
    const inside = root.contains(active)
    // wrap at the ends, and pull focus back in if it escaped the dialog
    if (e.shiftKey ? active === first || !inside : active === last || !inside) {
      e.preventDefault()
      ;(e.shiftKey ? last : first).focus()
    }
  }, [])

  return {
    backdropProps: { ref, onKeyDown },
    dialogProps: { role: 'dialog', 'aria-modal': true, 'aria-labelledby': titleId },
    titleProps: { id: titleId },
  }
}

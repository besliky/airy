import { beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { InsertPagesDialog } from '../src/renderer/InsertPagesDialog'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // LocaleProvider subscribes to the shell's language switch on mount
  Object.assign(window, { pdfApi: { onLanguageChanged: () => () => undefined } })
})

/**
 * UX-910: the insert-pages busy line is a live region that is MOUNTED BEFORE
 * the busy state flips — screen readers miss regions inserted into the DOM at
 * the moment of the change. Idle it must sit empty, busy it must speak.
 */

const source = {
  name: 'other.pdf',
  pages: [
    { width: 200, height: 300 },
    { width: 200, height: 300 },
  ],
}

function render(busy: boolean): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(InsertPagesDialog, {
          source,
          currentPage: 1,
          pos: 'end',
          onPos: () => undefined,
          afterPage: '1',
          onAfterPage: () => undefined,
          afterInvalid: false,
          range: '1-2',
          onRange: () => undefined,
          rangeInvalid: false,
          busy,
          onConfirm: () => undefined,
          onClose: () => undefined,
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

describe('InsertPagesDialog busy live region (UX-910)', () => {
  it('keeps the status region mounted and empty while idle', () => {
    const { container, unmount } = render(false)
    const region = container.querySelector('.pdf-modal-busy') as HTMLElement
    expect(region).not.toBeNull()
    expect(region.getAttribute('role')).toBe('status')
    expect(region.getAttribute('aria-live')).toBe('polite')
    expect(region.getAttribute('aria-atomic')).toBe('true')
    expect(region.textContent).toBe('')
    expect((container.querySelector('.pdf-modal') as HTMLElement).getAttribute('aria-busy')).toBe(
      'false',
    )
    unmount()
  })

  it('fills the same region when busy flips: spinner decorative, text spoken', () => {
    const { container, unmount } = render(true)
    const region = container.querySelector('.pdf-modal-busy') as HTMLElement
    expect(region.textContent).toBe('Inserting pages…')
    expect(region.querySelector('.pdf-modal-busy-spin')?.getAttribute('aria-hidden')).toBe('true')
    expect((container.querySelector('.pdf-modal') as HTMLElement).getAttribute('aria-busy')).toBe(
      'true',
    )
    unmount()
  })

  it('keeps every control disabled while busy (UX-706 regression guard)', () => {
    const { container, unmount } = render(true)
    const controls = [...container.querySelectorAll('input, button')]
    expect(controls.length).toBeGreaterThan(0)
    expect(controls.every((c) => (c as HTMLInputElement).disabled)).toBe(true)
    unmount()
  })
})

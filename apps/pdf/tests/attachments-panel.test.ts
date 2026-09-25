import { beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { AttachmentsPanel } from '../src/renderer/AttachmentsPanel'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // LocaleProvider subscribes to the shell's language switch on mount
  Object.assign(window, { pdfApi: { onLanguageChanged: () => () => undefined } })
})

/**
 * UX-1734: portfolio children were invisible in the UI — the panel makes them
 * visible (read-only list) and hands each row to the open action.
 */

const ATTACHMENTS = [
  { id: 'child-a.pdf', filename: 'child-a.pdf', description: '' },
  { id: 'child-b.pdf', filename: 'child-b.pdf', description: 'The second document' },
]

function render(
  attachments: typeof ATTACHMENTS,
  onOpen: (a: { id: string }) => void,
): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(AttachmentsPanel, {
          attachments,
          t: (key: string, params?: Record<string, unknown>) =>
            key === 'attachmentsCount'
              ? `${params?.n} attachments`
              : key === 'attachmentOpen'
                ? 'Open in tab'
                : key,
          onOpen,
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

describe('AttachmentsPanel (UX-1734)', () => {
  it('lists every embedded child with its name and an open affordance', () => {
    const { container, unmount } = render(ATTACHMENTS, () => undefined)
    const names = [...container.querySelectorAll('.pdf-attachment-name')].map(
      (el) => el.textContent,
    )
    expect(names).toEqual(['child-a.pdf', 'child-b.pdf'])
    expect(container.querySelectorAll('.pdf-attachment-open')).toHaveLength(2)
    // the description shows as the tooltip when present
    const second = container.querySelectorAll<HTMLElement>('.pdf-attachment-name')[1]!
    expect(second.getAttribute('data-tip')).toBe('The second document')
    unmount()
  })

  it('shows the count in the header and clicks through to onOpen', () => {
    const opened: string[] = []
    const { container, unmount } = render(ATTACHMENTS, (a) => opened.push(a.id))
    expect(container.querySelector('.pdf-attachments-title')?.textContent).toBe('2 attachments')
    const buttons = [...container.querySelectorAll<HTMLButtonElement>('.pdf-attachment-open')]
    act(() => {
      buttons[1]!.click()
    })
    expect(opened).toEqual(['child-b.pdf'])
    unmount()
  })
})

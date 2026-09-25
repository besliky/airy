// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { SaveCompatBanner } from '../src/renderer/save-compat-banner'
import type { SaveCompatFinding } from '../src/renderer/save-compat'

/// PAR-206: the pre-flight banner renders the fail-closed constructs found in
/// the workbook as a non-blocking role="status" strip, one item per finding
/// (plus its sheet/item detail), with a dismiss affordance. The
/// appear/disappear/dismiss policy itself is pinned in save-compat.test.ts
/// (saveCompatVisible); here we pin what the strip actually shows.

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  return loadLocale('en')
})

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

const FINDINGS: SaveCompatFinding[] = [
  { id: 'multi-select-dv', detail: 'Alpha' },
  { id: 'csv-flatten' },
]

describe('SaveCompatBanner', () => {
  it('renders a status strip with one list item per finding', () => {
    const { container, unmount } = render(
      createElement(SaveCompatBanner, { findings: FINDINGS, onDismiss: () => undefined }),
    )
    const banner = container.querySelector('.save-compat-banner')
    expect(banner?.getAttribute('role')).toBe('status')
    const title = container.querySelector('.save-compat-title')?.textContent
    expect(title).toBe("Some content in this workbook won't save to the file")
    const items = [...container.querySelectorAll('.save-compat-list li')]
    expect(items).toHaveLength(FINDINGS.length)
    // The multi-select item names its sheet; the CSV item needs no detail.
    expect(items[0]?.textContent).toContain('Multi-select list validation')
    expect(items[0]?.textContent).toContain('Alpha')
    expect(items[1]?.textContent).toContain('Saving as CSV keeps plain values')
    unmount()
  })

  it('localizes through the loaded dictionary (no raw keys leak)', () => {
    const { container, unmount } = render(
      createElement(SaveCompatBanner, { findings: FINDINGS, onDismiss: () => undefined }),
    )
    expect(container.textContent).not.toContain('appSaveCompat')
    unmount()
  })

  it('dismiss button fires onDismiss and is reachable by name', () => {
    const onDismiss = vi.fn()
    const { container, unmount } = render(
      createElement(SaveCompatBanner, { findings: FINDINGS, onDismiss }),
    )
    const button = container.querySelector('.save-compat-dismiss') as HTMLButtonElement
    expect(button.getAttribute('aria-label')).toBe('Dismiss the warning')
    act(() => {
      button.click()
    })
    expect(onDismiss).toHaveBeenCalledTimes(1)
    unmount()
  })
})

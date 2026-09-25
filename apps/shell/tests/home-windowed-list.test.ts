// PERF-1736, DOM level: the WindowedFileList must keep only the rows near the
// viewport in the DOM (plus spacers holding the scrolled-out space), and
// scrolling must swap the window. jsdom has no layout, so rect probes are
// stubbed to simulate a scrolling viewport; the window math itself is covered
// by home-window.test.ts.
/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
import {
  MIN_WINDOW_ROWS,
  OVERSCAN_ROWS,
  VIRTUALIZE_THRESHOLD,
} from '../src/renderer/src/home-window'
import { WindowedFileList } from '../src/renderer/src/home-window-list'

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

const ROWS = 2_000
const PITCH = 53
const VIEW_HEIGHT = 800

function rowEntry(i: number): RecentEntry {
  return {
    path: `/proj/p${String(i).padStart(5, '0')}.docx`,
    name: `p${String(i).padStart(5, '0')}.docx`,
    ext: 'docx',
    mtimeMs: 0,
    sizeBytes: 0,
    starred: false,
  }
}

const ROWS_2K = Array.from({ length: ROWS }, (_, i) => rowEntry(i))

function renderRow(entry: RecentEntry) {
  return createElement(
    'li',
    { className: 'recent-row', key: entry.path },
    createElement('span', { className: 'recent-name' }, entry.name),
  )
}

let host: HTMLDivElement
let root: Root

// rect-stub restore hooks, drained after every test
const rectCleanup: Array<() => void> = []

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  while (rectCleanup.length) rectCleanup.pop()!()
})

function rowNames(): string[] {
  return [...host.querySelectorAll('.recent-name')].map((n) => n.textContent)
}

function domRowCount(): number {
  return host.querySelectorAll('.recent-row').length
}

function spacerCount(): number {
  return host.querySelectorAll('.recent-list-spacer').length
}

interface RectLike {
  top: number
  bottom: number
  height: number
}

/**
 * Synthetic layout: the list (`ul.recent-list`) starts at the top of a
 * scroller whose viewport is VIEW_HEIGHT tall; `getScrollTop` tracks how far
 * the list is scrolled up. Every row measures one `pitch` tall.
 */
function stubRects(getScrollTop: () => number, pitch = PITCH): void {
  const proto = Element.prototype as unknown as {
    getBoundingClientRect: () => RectLike
  }
  const original = proto.getBoundingClientRect
  proto.getBoundingClientRect = function (this: Element) {
    const top = this.classList?.contains('recent-list') ? -getScrollTop() : 0
    const height = this.classList?.contains('scroll-viewport')
      ? VIEW_HEIGHT
      : this.classList?.contains('recent-list')
        ? ROWS * pitch
        : pitch
    return { top, bottom: top + height, height }
  }
  rectCleanup.push(() => {
    proto.getBoundingClientRect = original
  })
}

/** move the mounted list into a scrollable pane so findScrollParent finds it */
function mountInScrollPane(): HTMLDivElement {
  act(() => root.unmount())
  const pane = document.createElement('div')
  pane.className = 'scroll-viewport'
  // inline style so the jsdom cascade (the only cascade there is) reports it
  pane.style.overflowY = 'auto'
  host.append(pane)
  root = createRoot(pane)
  return pane
}

describe('WindowedFileList', () => {
  it('renders short lists in full, unchanged', () => {
    const rows = ROWS_2K.slice(0, VIRTUALIZE_THRESHOLD)
    act(() => {
      root.render(createElement(WindowedFileList, { rows, renderRow }))
    })
    expect(domRowCount()).toBe(VIRTUALIZE_THRESHOLD)
    expect(spacerCount()).toBe(0)
  })

  it('keeps only the initial window in the DOM for a huge list', () => {
    act(() => {
      root.render(createElement(WindowedFileList, { rows: ROWS_2K, renderRow }))
    })
    // unmeasured viewport → the MIN_WINDOW_ROWS fallback, never 2000 <li>
    expect(domRowCount()).toBe(MIN_WINDOW_ROWS)
    // rows at the top are the real ones; scrolled-out space is a spacer
    expect(rowNames()[0]).toBe('p00000.docx')
    expect(rowNames()).not.toContain('p01000.docx')
    expect(spacerCount()).toBe(1)
    const spacer = host.querySelector('.recent-list-spacer') as HTMLElement
    expect(spacer.style.height).toBe(`${(ROWS - MIN_WINDOW_ROWS) * PITCH}px`)
  })

  it('swaps the window when the page scroller scrolls to the bottom', async () => {
    let scrollTop = 0
    stubRects(() => scrollTop)
    const pane = mountInScrollPane()
    act(() => {
      root.render(createElement(WindowedFileList, { rows: ROWS_2K, renderRow }))
    })
    expect(rowNames()).toContain('p00000.docx')

    scrollTop = (ROWS - Math.floor(VIEW_HEIGHT / PITCH)) * PITCH
    await act(async () => {
      pane.dispatchEvent(new Event('scroll'))
      await Promise.resolve()
    })

    const names = rowNames()
    expect(names).toContain('p01999.docx')
    expect(names).not.toContain('p00000.docx')
    // the tail is fully rendered, so only the top spacer remains
    expect(spacerCount()).toBe(1)
    // DOM stays bounded after scrolling
    expect(domRowCount()).toBeLessThanOrEqual(MIN_WINDOW_ROWS + 2 * OVERSCAN_ROWS)
  })

  it('renders the middle window with spacers on both sides mid-scroll', async () => {
    let scrollTop = 0
    stubRects(() => scrollTop)
    const pane = mountInScrollPane()
    act(() => {
      root.render(createElement(WindowedFileList, { rows: ROWS_2K, renderRow }))
    })

    scrollTop = 1000 * PITCH
    await act(async () => {
      pane.dispatchEvent(new Event('scroll'))
      await Promise.resolve()
    })

    const names = rowNames()
    expect(names).toContain('p00995.docx')
    expect(names).not.toContain('p00000.docx')
    expect(names).not.toContain('p01999.docx')
    expect(spacerCount()).toBe(2)
  })

  it('follows the measured row pitch, not the built-in constant', async () => {
    let scrollTop = 0
    stubRects(() => scrollTop, 60)
    const pane = mountInScrollPane()
    act(() => {
      root.render(createElement(WindowedFileList, { rows: ROWS_2K, renderRow }))
    })

    scrollTop = 1000 * 60
    await act(async () => {
      pane.dispatchEvent(new Event('scroll'))
      await Promise.resolve()
    })
    // at a 60px pitch scrollTop 60000 is row ~1000; with the 53px default it
    // would land near row 1132 instead
    const names = rowNames()
    expect(names.some((n) => n !== null && n >= 'p00950.docx' && n <= 'p01050.docx')).toBe(true)
    expect(names.some((n) => n !== null && n > 'p01100.docx')).toBe(false)
  })
})

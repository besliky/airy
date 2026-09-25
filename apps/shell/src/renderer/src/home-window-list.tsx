import { useCallback, useEffect, useLayoutEffect, useRef, useState, type ReactElement } from 'react'
import type { RecentEntry } from '../../shared/home-api'
import {
  DEFAULT_ROW_PITCH,
  findScrollParent,
  spacerHeights,
  VIRTUALIZE_THRESHOLD,
  windowFor,
  type ListWindow,
} from './home-window'

/**
 * Windowed (virtualized) file list for the Home tables (PERF-1736). Above
 * VIRTUALIZE_THRESHOLD only the rows near the viewport (plus overscan) are
 * rendered, with two invisible spacer `<li>` elements holding the
 * scrolled-out space, so a 20k-file catalog keeps ~tens of `<li>` in the DOM
 * instead of 20 000 and every keystroke stays O(window) instead of O(corpus).
 *
 * The component deliberately does NOT own scrolling: the page scroller
 * (`.content`) keeps scrolling exactly as before, and the window tracks it
 * through rect maths (no layout-affecting CSS changes). At or below the
 * threshold it is a plain `<ul>` rendering every row — same DOM as before
 * windowing existed.
 */
export function WindowedFileList({
  rows,
  renderRow,
}: {
  rows: readonly RecentEntry[]
  renderRow: (entry: RecentEntry) => ReactElement
}): ReactElement {
  if (rows.length <= VIRTUALIZE_THRESHOLD) {
    return <ul className="recent-list">{rows.map((entry) => renderRow(entry))}</ul>
  }
  return <WindowedRows rows={rows} renderRow={renderRow} />
}

function WindowedRows({
  rows,
  renderRow,
}: {
  rows: readonly RecentEntry[]
  renderRow: (entry: RecentEntry) => ReactElement
}): ReactElement {
  const listRef = useRef<HTMLUListElement>(null)
  // the ancestor that scrolls this list (found once; does not change)
  const scrollerRef = useRef<HTMLElement | null>(null)
  // distance between consecutive row tops, measured from a rendered row
  const [pitch, setPitch] = useState(DEFAULT_ROW_PITCH)
  const [win, setWin] = useState<ListWindow>(() =>
    windowFor({ firstVisible: 0, visibleRows: 0, rowCount: rows.length }),
  )

  const measure = useCallback(() => {
    const list = listRef.current
    if (!list) return
    // measure the real pitch from a rendered row that carries its separator
    // border (every row except the very first) — falls back to the default
    // when nothing is measurable (jsdom, hidden window)
    const borderedRow = list.children.length > 1 ? (list.children[1] as HTMLElement) : null
    const measured = borderedRow ? borderedRow.getBoundingClientRect().height : 0
    if (measured > 0) setPitch((prev) => (Math.abs(prev - measured) > 0.5 ? measured : prev))
    const scroller = scrollerRef.current
    const listRect = list.getBoundingClientRect()
    let firstVisible = 0
    let visibleRows = 0
    if (scroller) {
      const view = scroller.getBoundingClientRect()
      const top = Math.max(listRect.top, view.top)
      const bottom = Math.min(listRect.bottom, view.bottom)
      firstVisible = Math.max(0, top - listRect.top) / pitch
      visibleRows = Math.max(0, bottom - top) / pitch
    }
    const next = windowFor({ firstVisible, visibleRows, rowCount: rows.length })
    setWin((prev) => (prev.start === next.start && prev.end === next.end ? prev : next))
  }, [pitch, rows.length])

  // locate the page scroller before the first measurement
  useLayoutEffect(() => {
    scrollerRef.current = findScrollParent(listRef.current)
  }, [])

  // re-measure after every commit that could move the window: mount, pitch
  // change, and corpus growth/shrink (progressive load, filter narrowing)
  useLayoutEffect(() => {
    measure()
  }, [measure])

  // the window tracks the page scroller and viewport resizes
  useEffect(() => {
    const scroller = scrollerRef.current
    if (!scroller) return
    scroller.addEventListener('scroll', measure, { passive: true })
    window.addEventListener('resize', measure)
    return () => {
      scroller.removeEventListener('scroll', measure)
      window.removeEventListener('resize', measure)
    }
  }, [measure])

  const { top, bottom } = spacerHeights(win, rows.length, pitch)
  return (
    <ul className="recent-list" ref={listRef}>
      {top > 0 && <li className="recent-list-spacer" style={{ height: top }} aria-hidden="true" />}
      {rows.slice(win.start, win.end).map((entry) => renderRow(entry))}
      {bottom > 0 && (
        <li className="recent-list-spacer" style={{ height: bottom }} aria-hidden="true" />
      )}
    </ul>
  )
}

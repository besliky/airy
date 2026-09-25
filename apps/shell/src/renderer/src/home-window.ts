/**
 * Windowing (list virtualization) math for the Home file tables (PERF-1736):
 * at the 20k-file catalog cap every keystroke re-rendered all 20 000 `<li>`
 * rows, freezing the tab for ~16s. Only the rows near the viewport are
 * rendered instead; everything here is pure arithmetic or a tiny DOM probe so
 * the rules stay unit-testable — the React side (home-window-list.tsx) only
 * measures and applies these numbers.
 */

/** rows rendered above and below the viewport so slow scrolls never show gaps */
export const OVERSCAN_ROWS = 8

/**
 * Fallback window size when the viewport cannot be measured (jsdom, hidden
 * window, zero-height scroller): small enough to bound the DOM, large enough
 * to cover a full screen at the default row pitch.
 */
export const MIN_WINDOW_ROWS = 48

/**
 * Windowing engages only above this many rows. Below it a full render is
 * cheaper than the windowing bookkeeping and keeps the whole list in the DOM
 * (find-in-page, tests); the 20k catalog cap is three orders of magnitude
 * larger, which is where rendering cost actually bites.
 */
export const VIRTUALIZE_THRESHOLD = 512

/**
 * Default distance between the tops of two consecutive rows, in px: the row
 * has 14+14px vertical padding around a 24px badge line, plus the 1px
 * separator border each row draws above itself (`.recent-list li + li`).
 * Only used until the real pitch is measured from a rendered row.
 */
export const DEFAULT_ROW_PITCH = 53

/** the rendered slice of a list; `end` is exclusive */
export interface ListWindow {
  start: number
  end: number
}

/**
 * Compute the visible slice. `firstVisible` is the fractional row index at
 * the viewport's top edge and `visibleRows` the fractional number of rows
 * that fit the viewport — both derived from rect measurements by the caller.
 * The window is clamped to the list and never smaller than MIN_WINDOW_ROWS
 * (bounded by the list length) when the viewport is unmeasurable.
 */
export function windowFor(params: {
  firstVisible: number
  visibleRows: number
  rowCount: number
  overscan?: number
  minWindow?: number
}): ListWindow {
  const overscan = params.overscan ?? OVERSCAN_ROWS
  const minWindow = params.minWindow ?? MIN_WINDOW_ROWS
  const { rowCount } = params
  const windowRows = Math.max(
    Math.ceil(Math.max(0, params.visibleRows)) + 2 * overscan,
    Math.min(minWindow, rowCount),
  )
  const start = Math.min(
    Math.max(0, Math.floor(Math.max(0, params.firstVisible)) - overscan),
    Math.max(0, rowCount),
  )
  const end = Math.min(rowCount, start + windowRows)
  return { start, end }
}

/**
 * Heights of the invisible spacer `<li>` elements that hold the scrolled-out
 * space above and below the rendered window, so the scroller's total
 * scrollable height matches an unvirtualized list.
 */
export function spacerHeights(
  win: ListWindow,
  rowCount: number,
  rowPitch: number,
): { top: number; bottom: number } {
  return { top: win.start * rowPitch, bottom: (rowCount - win.end) * rowPitch }
}

/**
 * Nearest ancestor that actually scrolls (overflow auto/scroll/overlay), so
 * the window tracks the page scroller (`.content`) instead of owning a
 * scrollbar of its own. Falls back to the document scroller — including in
 * environments with no cascade (jsdom), where every probe reads as static.
 */
export function findScrollParent(el: HTMLElement | null): HTMLElement | null {
  let node = el?.parentElement ?? null
  while (node) {
    const overflowY = getComputedStyle(node).overflowY
    if (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay') return node
    node = node.parentElement
  }
  return (document.scrollingElement as HTMLElement | null) ?? document.documentElement
}

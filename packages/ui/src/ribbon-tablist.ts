/**
 * WAI-ARIA tabs pattern for the ribbon tab strip: the strip becomes a
 * `tablist` whose tabs use a roving tabindex (only the selected tab is in
 * the page tab order) with automatic activation — Left/Right (mirrored in
 * RTL) move focus and switch tabs, Home/End jump to the strip's ends.
 * The command band becomes the `tabpanel` labelled by the selected tab.
 *
 * Only key presses that bubble up from the tab buttons themselves are
 * handled, so keys the editors bind inside the band are never stolen.
 */
import { useCallback, useRef, type KeyboardEvent, type RefCallback } from 'react'

/** Wrap `current` by `delta` within `[0, count)` (APG tabs wrap at the ends). */
export function wrapRibbonTabIndex(current: number, count: number, delta: 1 | -1): number {
  if (count <= 0) return -1
  return (((current + delta) % count) + count) % count
}

/**
 * Index the tablist key navigates to from `current`, or -1 when the key is
 * not tabstrip navigation (caller leaves it for the page). `rtl` mirrors
 * the horizontal arrows for right-to-left chrome.
 */
export function nextRibbonTabIndex(
  current: number,
  count: number,
  key: string,
  rtl = false,
): number {
  if (count <= 0) return -1
  switch (key) {
    case 'ArrowRight':
      return wrapRibbonTabIndex(current, count, rtl ? -1 : 1)
    case 'ArrowLeft':
      return wrapRibbonTabIndex(current, count, rtl ? 1 : -1)
    case 'Home':
      return 0
    case 'End':
      return count - 1
    default:
      return -1
  }
}

export function ribbonTabId(idPrefix: string, tab: string): string {
  // tab ids come from app state and may contain spaces ("Page Layout") or
  // other characters HTML ids must not (ASCII whitespace)
  return `${idPrefix}-tab-${tab.replace(/[^\w-]+/g, '-')}`
}

/** Props for the command band element, labelled by the selected tab button. */
export function ribbonPanelProps(idPrefix: string, activeTab: string | null) {
  return {
    role: 'tabpanel' as const,
    id: `${idPrefix}-panel`,
    'aria-labelledby': activeTab === null ? undefined : ribbonTabId(idPrefix, activeTab),
  }
}

export interface UseRibbonTablistOptions {
  /** Visible tab ids in strip order (contextual tabs appended when shown). */
  readonly tabs: readonly string[]
  /** Currently selected tab id, or null when nothing is selected yet. */
  readonly activeTab: string | null
  /** Stable DOM id prefix for this app's ribbon (`<prefix>-tab-<id>`). */
  readonly idPrefix: string
  /** Localized name announced for the tablist ("Ribbon tabs"). */
  readonly label: string
  /** Full tab selection (mirrors the click handler, including peek logic). */
  readonly onSelect: (tab: string) => void
}

export interface RibbonTabProps {
  readonly role: 'tab'
  readonly id: string
  readonly 'aria-selected': boolean
  readonly tabIndex: number
  readonly ref: RefCallback<HTMLButtonElement | null>
}

export interface RibbonTablistController {
  /** Spread on the strip wrapper (a `display: contents` div around the tabs). */
  readonly tablistProps: {
    readonly role: 'tablist'
    readonly 'aria-label': string
    readonly onKeyDown: (e: KeyboardEvent<HTMLElement>) => void
  }
  /** Spread on each tab button next to its className/key/onClick. */
  readonly tabProps: (tab: string) => RibbonTabProps
}

export function useRibbonTablist(options: UseRibbonTablistOptions): RibbonTablistController {
  const { tabs, activeTab, idPrefix, label, onSelect } = options
  const buttons = useRef(new Map<string, HTMLButtonElement>())
  const activeIndex = tabs.indexOf(activeTab ?? '')
  // roving tabindex: the selected tab (or the first tab before any selection)
  const focusIndex = activeIndex >= 0 ? activeIndex : 0

  const tablistProps = {
    role: 'tablist' as const,
    'aria-label': label,
    onKeyDown: useCallback(
      (e: KeyboardEvent<HTMLElement>) => {
        if (e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return
        const rtl = document.documentElement.dir === 'rtl'
        const next = nextRibbonTabIndex(focusIndex, tabs.length, e.key, rtl)
        if (next < 0) return
        e.preventDefault()
        const tab = tabs[next]
        if (tab === undefined) return
        onSelect(tab)
        // the button node survives the re-render (stable keys), so focusing
        // right away lands on the newly selected tab
        buttons.current.get(tab)?.focus()
      },
      [focusIndex, tabs, onSelect],
    ),
  }

  const tabProps = useCallback(
    (tab: string): RibbonTabProps => ({
      role: 'tab',
      id: ribbonTabId(idPrefix, tab),
      'aria-selected': tab === activeTab,
      tabIndex: tab === activeTab || (activeTab === null && tabs[0] === tab) ? 0 : -1,
      ref: (el) => {
        if (el) buttons.current.set(tab, el)
        else buttons.current.delete(tab)
      },
    }),
    [activeTab, idPrefix, tabs],
  )

  return { tablistProps, tabProps }
}

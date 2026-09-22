/**
 * BUG-1531: the Recents list on Home only reloaded on window focus and view or
 * filter switches. Opening files from quick-cards or session restore writes
 * recents on disk, but returning to the Home tab raised no focus event — the
 * section kept showing "No recent files" until the next alt-tab. The fix
 * reloads the list from the tabs-changed broadcast whenever it leaves Home
 * the visible tab; this test drives that subscription with fake tab events.
 */
/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { TabSummary } from '../src/shared/tabs-api'
import { Home } from '../src/renderer/src/Home'
import { LocaleProvider } from '../src/renderer/src/locale'

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

// jsdom has no IntersectionObserver; Home only builds one when a page has a
// load-more sentinel, but a stub keeps the test independent of that detail
class IntersectionObserverStub implements IntersectionObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
  takeRecords(): IntersectionObserverEntry[] {
    return []
  }
  readonly root = null
  readonly rootMargin = ''
  readonly thresholds = []
}
;(globalThis as typeof globalThis & Record<string, unknown>).IntersectionObserver ??=
  IntersectionObserverStub

const EMPTY_PAGE = { entries: [], total: 0, totalAll: 0 }

let host: HTMLDivElement
let root: Root
let recents: ReturnType<typeof vi.fn>
let changedHandlers: Array<(tabs: TabSummary[]) => void>

const tab = (id: string, kind: TabSummary['kind'], active: boolean): TabSummary => ({
  id,
  kind,
  title: id,
  closable: id !== 'home',
  active,
})

beforeEach(() => {
  vi.useFakeTimers()
  changedHandlers = []
  recents = vi.fn(async () => EMPTY_PAGE)
  window.aiOffice = {
    recents,
    starred: vi.fn(async () => EMPTY_PAGE),
  } as unknown as typeof window.aiOffice
  window.aiOfficeTabs = {
    list: vi.fn(async () => [] as TabSummary[]),
    onChanged: vi.fn((handler: (tabs: TabSummary[]) => void) => {
      changedHandlers.push(handler)
      return () => {}
    }),
  } as unknown as typeof window.aiOfficeTabs
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.useRealTimers()
  delete (window as Partial<typeof window>).aiOfficeTabs
})

async function renderHome(): Promise<void> {
  await act(async () => {
    root.render(createElement(LocaleProvider, { initial: 'en' }, createElement(Home)))
  })
}

function broadcast(tabs: TabSummary[]): void {
  act(() => {
    for (const handler of changedHandlers) handler(tabs)
  })
}

/** advance the debounce and let the reload's promise callbacks settle inside act */
async function settle(ms: number): Promise<void> {
  await act(async () => {
    vi.advanceTimersByTime(ms)
    await Promise.resolve()
  })
}

describe('home recents live refresh (BUG-1531)', () => {
  it('reloads the list when a tabs-changed broadcast leaves Home active', async () => {
    await renderHome()
    expect(recents).toHaveBeenCalledTimes(1)

    // a document tab opened and activated: Home is hidden, nothing to do yet
    broadcast([tab('home', 'home', false), tab('t1', 'docs', true)])
    await settle(200)
    expect(recents).toHaveBeenCalledTimes(1)

    // the user clicks back on the Home tab: the broadcast now names Home
    // active, and the debounced reload re-pulls the recents page
    broadcast([tab('home', 'home', true), tab('t1', 'docs', false)])
    expect(recents).toHaveBeenCalledTimes(1)
    await settle(150)
    expect(recents).toHaveBeenCalledTimes(2)
    expect(recents).toHaveBeenLastCalledWith({ offset: 0, limit: 50, ext: undefined })
  })

  it('collapses a burst of broadcasts into one reload', async () => {
    await renderHome()
    expect(recents).toHaveBeenCalledTimes(1)

    broadcast([tab('home', 'home', true), tab('t1', 'pdf', false)])
    broadcast([tab('home', 'home', true), tab('t2', 'markdown', false)])
    broadcast([tab('home', 'home', true), tab('t3', 'html', false)])
    await settle(150)
    expect(recents).toHaveBeenCalledTimes(2)
  })

  it('ignores broadcasts while another tab is the visible one', async () => {
    await renderHome()
    broadcast([tab('home', 'home', false), tab('t1', 'sheets', true)])
    await settle(500)
    expect(recents).toHaveBeenCalledTimes(1)
  })
})

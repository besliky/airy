// TEST-704: the Home file-search wiring around the pure helpers in
// home-search.ts had no coverage — only `filterFileEntries` & friends were
// tested. These DOM tests drive the real Home component: the live filtering
// of the visible rows, the localized no-match empty state, the clear button,
// Escape (clear + blur, with the IME-composition guard) and the ⌘/Ctrl+F
// focus shortcut. TEST-704 also names an "aria-live count" of matches — no
// such element exists in Home today, so that part is documented as a
// non-applicable prod change, not tested.
/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
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

const entry = (path: string, name: string, ext: string): RecentEntry => ({
  path,
  name,
  ext,
  mtimeMs: 1_700_000_000_000,
  sizeBytes: 128,
  starred: false,
})

const FILES = [
  entry('/docs/Report Q3.docx', 'Report Q3.docx', 'docx'),
  entry('/sheets/budget.xlsx', 'budget.xlsx', 'xlsx'),
  entry('/notes/ideas.md', 'ideas.md', 'md'),
]

let host: HTMLDivElement
let root: Root

beforeEach(() => {
  const page = { entries: FILES, total: FILES.length, totalAll: FILES.length }
  window.aiOffice = {
    recents: vi.fn(async () => page),
    starred: vi.fn(async () => page),
  } as unknown as typeof window.aiOffice
  window.aiOfficeTabs = {
    list: vi.fn(async () => [] as TabSummary[]),
    onChanged: vi.fn(() => () => {}),
  } as unknown as typeof window.aiOfficeTabs
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  delete (window as Partial<typeof window>).aiOfficeTabs
})

async function renderHome(): Promise<void> {
  await act(async () => {
    root.render(createElement(LocaleProvider, { initial: 'en' }, createElement(Home)))
    // flush the initial recents/starred loads
    await Promise.resolve()
  })
}

function searchInput(): HTMLInputElement {
  return document.querySelector('.file-search-input') as HTMLInputElement
}

/** React-controlled input: set the value through the native setter so onChange fires */
function typeQuery(input: HTMLInputElement, value: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  act(() => {
    setter.call(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function visibleNames(): string[] {
  return [...document.querySelectorAll('.recent-name')].map((n) => n.textContent)
}

describe('home file search wiring (TEST-704)', () => {
  it('filters the visible rows live and shows the localized empty state', async () => {
    await renderHome()
    expect(visibleNames()).toEqual(['Report Q3.docx', 'budget.xlsx', 'ideas.md'])

    typeQuery(searchInput(), 'BUDGET')
    expect(visibleNames()).toEqual(['budget.xlsx'])
    // the filter is a file-NAME match: a path fragment alone matches nothing
    typeQuery(searchInput(), '/docs')
    expect(visibleNames()).toEqual([])
    expect(document.querySelector('.empty-hint')!.textContent).toBe('No files matching "/docs".')

    // clearing restores the unfiltered list without any refetch of recents
    typeQuery(searchInput(), '')
    expect(visibleNames()).toEqual(['Report Q3.docx', 'budget.xlsx', 'ideas.md'])
  })

  it('offers a clear button while a query is set, and it resets the filter', async () => {
    await renderHome()
    expect(document.querySelector('.file-search-clear')).toBeNull()
    typeQuery(searchInput(), 'report')
    const clear = document.querySelector('.file-search-clear') as HTMLButtonElement
    expect(clear.getAttribute('aria-label')).toBe('Clear search')
    act(() => clear.click())
    expect(searchInput().value).toBe('')
    expect(visibleNames()).toEqual(['Report Q3.docx', 'budget.xlsx', 'ideas.md'])
  })

  it('Escape clears the field and blurs it, unless an IME composition holds it', async () => {
    await renderHome()
    const input = searchInput()
    typeQuery(input, 'budget')
    input.focus()
    const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })
    act(() => input.dispatchEvent(escape))
    expect(input.value).toBe('')
    expect(document.activeElement).not.toBe(input)

    // pinyin IME: Escape first cancels the composition — it must not wipe
    // the field (the guard reads nativeEvent.isComposing)
    typeQuery(input, 'nihao')
    input.focus()
    const composing = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })
    Object.defineProperty(composing, 'isComposing', { value: true })
    act(() => input.dispatchEvent(composing))
    expect(input.value).toBe('nihao')
  })

  it('⌘/Ctrl+F focuses and selects the search field', async () => {
    await renderHome()
    const input = searchInput()
    typeQuery(input, 'budget')
    expect(document.activeElement).not.toBe(input)
    const ctrlF = new KeyboardEvent('keydown', {
      key: 'f',
      code: 'KeyF',
      ctrlKey: true,
      bubbles: true,
      cancelable: true,
    })
    act(() => window.dispatchEvent(ctrlF))
    expect(document.activeElement).toBe(input)
    // the whole query arrives preselected, ready to be replaced
    expect(input.selectionStart).toBe(0)
    expect(input.selectionEnd).toBe(input.value.length)
    expect(ctrlF.defaultPrevented).toBe(true)
  })
})

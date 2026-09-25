// PERF-1736 regression, DOM level: a full 20k-file project (the catalog cap)
// must stay interactive in the Home window. Drives the real Home component
// against stubbed IPC: the mounted list keeps only a windowed slice of <li>
// (not 20 000), the #179 needle (file #280) is found by live search, the
// counter stays honest, and a measured keystroke over the whole corpus stays
// within the 200ms budget (median over the query's characters).
/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { ProjectHomeApi, ProjectSummaryEntry, RecentEntry } from '../src/shared/home-api'
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

const FILES = 20_000

function projectPaths(count: number): string[] {
  return Array.from({ length: count }, (_, i) => `/proj/p${String(i).padStart(5, '0')}.docx`)
}

function entryFor(path: string): RecentEntry {
  return {
    path,
    name: path.split('/').pop() as string,
    ext: 'docx',
    mtimeMs: 1_700_000_000_000,
    sizeBytes: 128,
    starred: false,
  }
}

const PROJECT: ProjectSummaryEntry = {
  id: 'p1',
  name: 'Huge project',
  createdAt: '2026-01-01T00:00:00.000Z',
  updatedAt: '2026-01-01T00:00:00.000Z',
  fileCount: FILES,
  lastActiveAt: '2026-01-01T00:00:00.000Z',
  isDefault: false,
}

let host: HTMLDivElement
let root: Root

beforeEach(() => {
  const page = { entries: [], total: 0, totalAll: 0 }
  window.aiOffice = {
    recents: vi.fn(async () => page),
    starred: vi.fn(async () => page),
    statPaths: vi.fn(async (chunk: string[]) => chunk.map(entryFor)),
  } as unknown as typeof window.aiOffice
  window.aiOfficeProject = {
    listProjects: vi.fn(async () => [PROJECT]),
    listFiles: vi.fn(async () => projectPaths(FILES)),
  } as unknown as ProjectHomeApi
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
  delete (window as Partial<typeof window>).aiOfficeProject
  delete (window as Partial<typeof window>).aiOfficeTabs
})

async function renderHome(): Promise<void> {
  await act(async () => {
    root.render(createElement(LocaleProvider, { initial: 'en' }, createElement(Home)))
    await Promise.resolve()
  })
}

/** select the project, then let the chunked 20k-catalog load settle */
async function openProject(): Promise<void> {
  const item = document.querySelector('.proj-item-main') as HTMLDivElement
  await act(async () => {
    item.click()
  })
  // poll until the progressive chunked load has stopped changing the list
  let prev = -1
  let stable = 0
  for (let i = 0; i < 400; i++) {
    await new Promise((resolve) => setTimeout(resolve, 10))
    const count = document.querySelectorAll('.recent-row').length
    if (count > 0 && count === prev) {
      if (++stable >= 5) return
    } else {
      stable = 0
      prev = count
    }
  }
  throw new Error('20k catalog never settled')
}

function fileCount(): string {
  return (document.querySelector('.file-count') as HTMLElement).textContent ?? ''
}

function visibleNames(): string[] {
  return [...document.querySelectorAll('.recent-name')].map((n) => n.textContent)
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

describe('20k-file project stays interactive (PERF-1736)', () => {
  it('mounts the loaded catalog as a bounded window of <li>, not 20 000', async () => {
    await renderHome()
    await openProject()
    const liCount = document.querySelectorAll('.recent-row').length
    expect(liCount).toBeGreaterThan(0)
    // a window (plus overscan and spacers), never the whole catalog
    expect(liCount).toBeLessThan(200)
    // honest counter for the full (loaded) catalog
    expect(fileCount()).toBe('20000 files')
    expect(visibleNames()).toContain('p00000.docx')
  }, 30_000)

  it('finds file #280 by search and shows the single match', async () => {
    await renderHome()
    await openProject()
    typeQuery(searchInput(), 'p00280')
    expect(visibleNames()).toEqual(['p00280.docx'])
    expect(fileCount()).toBe('1 file')
    expect(document.querySelector('.empty-hint')).toBeNull()
  }, 30_000)

  it('keeps each keystroke within the 200ms budget (median, 20k corpus)', async () => {
    await renderHome()
    await openProject()
    const input = searchInput()
    // rare needle: every prefix must sweep the whole 20k corpus and re-render
    const needle = 'p19876'
    const perKeystroke: number[] = []
    for (let i = 1; i <= needle.length; i++) {
      const t0 = performance.now()
      typeQuery(input, needle.slice(0, i))
      perKeystroke.push(performance.now() - t0)
    }
    const sorted = [...perKeystroke].sort((a, b) => a - b)
    const median = sorted[Math.floor(sorted.length / 2)] as number
    expect(visibleNames()).toEqual(['p19876.docx'])
    // 200ms, not 100ms: CI runners are shared and load-sensitive (evening
    // medians up to 127ms observed on byte-identical code), while the
    // regression this guards against is ~61,000ms per keystroke, so 200ms
    // keeps ~300x sensitivity; local medians are 5-32ms.
    expect(median).toBeLessThan(200)
    // and the worst keystroke stays in interaction territory, not O(corpus)
    expect(perKeystroke[perKeystroke.length - 1] as number).toBeLessThan(1_000)
  }, 30_000)

  it('clearing the search restores the bounded window', async () => {
    await renderHome()
    await openProject()
    typeQuery(searchInput(), 'p00280')
    expect(visibleNames()).toEqual(['p00280.docx'])
    typeQuery(searchInput(), '')
    expect(document.querySelectorAll('.recent-row').length).toBeLessThan(200)
    expect(fileCount()).toBe('20000 files')
  }, 30_000)
})

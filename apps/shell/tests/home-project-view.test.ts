// BUG-1676 regression, DOM level: the project files view must index (and let
// search find) the whole catalog, not silently stop at the old 256-path cap.
// Drives the real Home component against stubbed IPC: a 300-file project is
// fully listed with an honest counter, the old-cap victim p0280 is found by
// the live search, and the stat round-trips arrive in bounded chunks.
/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
import { STAT_CHUNK_SIZE } from '../src/renderer/src/project-files'
import { Home } from '../src/renderer/src/Home'
import { LocaleProvider } from '../src/renderer/src/locale'
import type { ProjectSummaryEntry, ProjectHomeApi } from '../src/shared/home-api'
import type { TabSummary } from '../src/shared/tabs-api'

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

const FILES = 300

function projectPaths(count: number): string[] {
  return Array.from({ length: count }, (_, i) => `/proj/p${String(i).padStart(4, '0')}.docx`)
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
  name: 'Big project',
  createdAt: '2026-01-01T00:00:00.000Z',
  updatedAt: '2026-01-01T00:00:00.000Z',
  fileCount: FILES,
  lastActiveAt: '2026-01-01T00:00:00.000Z',
  isDefault: false,
}

let host: HTMLDivElement
let root: Root
let statChunkSizes: number[]

beforeEach(() => {
  const page = { entries: [], total: 0, totalAll: 0 }
  statChunkSizes = []
  window.aiOffice = {
    recents: vi.fn(async () => page),
    starred: vi.fn(async () => page),
    // in-memory stat: records the chunking the loader actually performs
    statPaths: vi.fn(async (chunk: string[]) => {
      statChunkSizes.push(chunk.length)
      return chunk.map(entryFor)
    }),
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
    // flush the initial recents/starred/project loads
    await Promise.resolve()
  })
}

/** select the project, then let the chunked catalog load settle */
async function openProject(): Promise<void> {
  const item = document.querySelector('.proj-item-main') as HTMLDivElement
  await act(async () => {
    item.click()
    // macrotask flush: chunked loading yields between chunks
    await new Promise((resolve) => setTimeout(resolve, 20))
  })
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

describe('project files view indexing (BUG-1676)', () => {
  it('lists a 300-file project fully with an honest counter', async () => {
    await renderHome()
    await openProject()
    // the old cap truncated this view at "256 files"
    expect(fileCount()).toBe('300 files')
    expect(visibleNames()).toHaveLength(300)
    expect(visibleNames()).toContain('p0280.docx')
  })

  it('stats the catalog in bounded chunks, not one giant round-trip', async () => {
    await renderHome()
    await openProject()
    expect(statChunkSizes.length).toBeGreaterThan(0)
    for (const size of statChunkSizes) expect(size).toBeLessThanOrEqual(STAT_CHUNK_SIZE)
    expect(statChunkSizes.reduce((a, b) => a + b, 0)).toBe(300)
  })

  it('finds by search the file that used to sit beyond the old cap', async () => {
    await renderHome()
    await openProject()
    typeQuery(searchInput(), 'p0280')
    expect(visibleNames()).toEqual(['p0280.docx'])
    // counter narrows to the single match, no false "no files" state
    expect(fileCount()).toBe('1 file')
    expect(document.querySelector('.empty-hint')).toBeNull()
  })
})

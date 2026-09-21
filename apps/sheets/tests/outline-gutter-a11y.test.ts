// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from 'vitest'

import { IRenderManagerService } from '@univerjs/engine-render'

import { installOutlineGutter } from '../src/renderer/outline-gutter'
import { loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

/**
 * UX-1102: the outline gutter's +/- buttons are real focusable <button>s
 * appended to <body>; they must also be operable blind — named by axis and
 * level (not their +/− glyph), carrying the group's expanded state, and
 * keeping keyboard focus inside the gutter when a focused group scrolls
 * off-screen. This mounts the DOM overlay against a stub Univer runtime
 * (the gutter only reads viewport scroll + cell rects + the lazy state).
 */
beforeAll(() => {
  return loadLocale('en')
})

setModuleLang('en')

const CELL = { width: 100, height: 20 }

function stubGeometry(main: DOMRect, rowStrip: DOMRect | null, colStrip: DOMRect | null): void {
  // each call re-seeds the whole geometry: the gutter resolves the host and
  // the strips via getElementById/querySelectorAll, so stale hosts from a
  // previous test would shadow the new canvases
  for (const stale of document.querySelectorAll('#univer-container')) stale.remove()
  const rect = (r: DOMRect | null) => () => (r ?? new DOMRect(0, -9999, 0, 0)) as DOMRect
  const canvases = [
    { el: document.createElement('canvas'), r: main },
    { el: document.createElement('canvas'), r: rowStrip },
    { el: document.createElement('canvas'), r: colStrip },
  ]
  const host = document.createElement('div')
  host.id = 'univer-container'
  for (const { el, r } of canvases) {
    el.getBoundingClientRect = rect(r)
    host.append(el)
  }
  document.body.append(host)
}

/** a workbook whose sheet has rows 1-2 grouped at level 1 (summary below) */
function stubRuntime(rows: Map<number, { level: number; collapsed: boolean }>) {
  const state = {
    file: { sheets: [{ id: 'sh', rowCount: 10, columnCount: 10 }] },
    editJournal: { structuralOps: new Map() },
    outline: new Map([['sh', { rows, cols: new Map() }]]),
  } as unknown as LazyWorkbookState
  const viewport = {
    viewportScrollX: 0,
    viewportScrollY: 0,
    onScrollAfter$: { subscribe: () => ({ unsubscribe: () => undefined }) },
  }
  const renderManager = {
    getRenderById: () => ({ scene: { getViewport: () => viewport } }),
  }
  const injector = {
    get: (token: unknown) =>
      token === IRenderManagerService
        ? renderManager
        : { onCommandExecuted: () => ({ dispose: () => undefined }) },
  }
  const worksheet = {
    getSheetId: () => 'sh',
    getZoom: () => 1,
    getRowHeight: () => CELL.height,
    getColumnWidth: () => CELL.width,
    getMaxRows: () => 10,
    getMaxColumns: () => 10,
    getRange: (row: number, column: number) => ({
      getCellRect: () => ({ x: 100 + column * CELL.width, y: 50 + row * CELL.height }),
    }),
  }
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({ getId: () => 'wb', getActiveSheet: () => worksheet }),
    },
    univer: { __getInjector: () => injector },
  }
  return { runtime, lazyWorkbookRef: { current: state } }
}

const flushed = async (handle: { refresh(): void } | null): Promise<void> => {
  // update() runs inside requestAnimationFrame
  await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)))
  handle?.refresh()
  await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)))
}

/** rows 1-2 grouped at level 1, summary below at row 3 (collapsed flag lives on the summary line) */
const groupAt = (collapsed: boolean): Map<number, { level: number; collapsed: boolean }> =>
  new Map([
    [1, { level: 1, collapsed: false }],
    [2, { level: 1, collapsed: false }],
    [3, { level: 0, collapsed }],
  ])

describe('outline gutter accessibility (UX-1102)', () => {
  it('names the buttons by axis+level and exposes the group state', async () => {
    stubGeometry(
      new DOMRect(100, 50, 800, 600),
      new DOMRect(80, 50, 20, 600),
      new DOMRect(100, 20, 800, 30),
    )
    const { runtime, lazyWorkbookRef } = stubRuntime(groupAt(false))
    const onToggle = vi.fn()
    const handle = installOutlineGutter(runtime as never, lazyWorkbookRef, {
      placement: () => ({ summaryBelow: true, summaryRight: true }),
      onToggle,
    })
    expect(handle).not.toBeNull()
    await flushed(handle)

    // rows 1-2 at level 1, summary below at row 3 → exactly one button
    const button = document.querySelector('.outline-gutter-button') as HTMLButtonElement
    expect(button).not.toBeNull()
    expect(button.type).toBe('button')
    expect(button.textContent).toBe('−')
    expect(button.getAttribute('aria-expanded')).toBe('true')
    expect(button.getAttribute('aria-label')).toBe(t('appOutlineGutterCollapseRows', { n: 1 }))

    // the collapsed twin flips both the label and the state
    lazyWorkbookRef.current = {
      ...lazyWorkbookRef.current,
      outline: new Map([['sh', { rows: groupAt(true), cols: new Map() }]]),
    } as unknown as LazyWorkbookState
    await flushed(handle)
    expect(button.textContent).toBe('+')
    expect(button.getAttribute('aria-expanded')).toBe('false')
    expect(button.getAttribute('aria-label')).toBe(t('appOutlineGutterExpandRows', { n: 1 }))
    handle?.dispose()
  })

  it('hands focus to the surviving neighbour when the focused button goes away', async () => {
    stubGeometry(
      new DOMRect(100, 50, 800, 600),
      new DOMRect(80, 50, 20, 600),
      new DOMRect(100, 20, 800, 30),
    )
    // two groups: rows 1-2 (summary 3) and rows 5-6 (summary 7)
    const rows = new Map([
      [1, { level: 1, collapsed: false }],
      [2, { level: 1, collapsed: false }],
      [3, { level: 0, collapsed: false }],
      [5, { level: 1, collapsed: false }],
      [6, { level: 1, collapsed: false }],
      [7, { level: 0, collapsed: false }],
    ])
    const { runtime, lazyWorkbookRef } = stubRuntime(rows)
    const handle = installOutlineGutter(runtime as never, lazyWorkbookRef, {
      placement: () => ({ summaryBelow: true, summaryRight: true }),
      onToggle: vi.fn(),
    })
    expect(handle).not.toBeNull()
    await flushed(handle)
    const buttons = [...document.querySelectorAll('.outline-gutter-button')] as HTMLButtonElement[]
    expect(buttons.length).toBe(2)

    // the second group scrolls out of view: its button is not wanted anymore
    const gone = buttons[1]!
    const survive = buttons[0]!
    gone.focus()
    lazyWorkbookRef.current = {
      ...lazyWorkbookRef.current,
      outline: new Map([['sh', { rows: groupAt(false), cols: new Map() }]]),
    } as unknown as LazyWorkbookState
    await flushed(handle)
    expect(document.body.contains(gone)).toBe(false)
    expect(document.activeElement).toBe(survive)
    handle?.dispose()
  })
})

describe('outline gutter stays on the header strip (UX-1106)', () => {
  /**
   * The gutter is a fixed layer above the app chrome; level lanes reaching
   * past the strip's outer edge drew the level 2+ buttons over the formula
   * bar / Name Box. Every button must sit inside its strip; a lane that
   * does not fit clamps to the strip edge (keeping its own summary line,
   * so clamped buttons never stack).
   */
  it('clamps deep row lanes to the strip edge instead of the formula bar', async () => {
    // row header strip 46 px wide (surface at x=100): lanes up to level 3
    // fit (100-3-42 = 55), level 4 (41) would land on the Name Box
    stubGeometry(
      new DOMRect(100, 50, 800, 600),
      new DOMRect(54, 50, 46, 600),
      new DOMRect(100, 20, 800, 30),
    )
    // rows 1-2 grouped at level 1 (summary 3), rows 5-6 at level 4
    // (summary 7): the level-4 run nests — it draws one button per level
    // 1..4, only the deepest one needs the clamp
    const rows = new Map([
      [1, { level: 1, collapsed: false }],
      [2, { level: 1, collapsed: false }],
      [3, { level: 0, collapsed: false }],
      [5, { level: 4, collapsed: false }],
      [6, { level: 4, collapsed: false }],
      [7, { level: 0, collapsed: false }],
    ])
    const { runtime, lazyWorkbookRef } = stubRuntime(rows)
    const handle = installOutlineGutter(runtime as never, lazyWorkbookRef, {
      placement: () => ({ summaryBelow: true, summaryRight: true }),
      onToggle: vi.fn(),
    })
    expect(handle).not.toBeNull()
    await flushed(handle)
    const buttons = [...document.querySelectorAll('.outline-gutter-button')] as HTMLButtonElement[]
    // the level-1 run plus the level-4 run's four nested levels
    expect(buttons.length).toBe(5)
    for (const button of buttons) {
      // never past the strip's outer edge onto the Name Box / formula bar
      expect(Number.parseFloat(button.style.left)).toBeGreaterThanOrEqual(54)
    }
    // in-lane buttons keep their own lane (level 1 = 83, level 2 = 69,
    // level 3 = 55)…
    const lefts = buttons.map((button) => button.style.left).sort()
    expect(lefts).toEqual(['54px', '55px', '69px', '83px', '83px'])
    // …and the level-4 button (aria-label carries its level) is the clamped one
    const deepest = buttons.find((button) =>
      button.getAttribute('aria-label')?.includes(t('appOutlineGutterCollapseRows', { n: 4 })),
    )
    expect(deepest?.style.left).toBe('54px')
    handle?.dispose()
  })

  it('clamps deep column lanes to the strip edge (level 2 already overflows a 30 px strip)', async () => {
    stubGeometry(
      new DOMRect(100, 50, 800, 600),
      new DOMRect(80, 50, 20, 600),
      new DOMRect(100, 20, 800, 30),
    )
    // columns 0-1 grouped at level 1 (summary 2), columns 3-4 at level 2
    // (summary 5): the nested level-2 run draws levels 1 and 2
    const cols = new Map([
      [0, { level: 1, collapsed: false }],
      [1, { level: 1, collapsed: false }],
      [2, { level: 0, collapsed: false }],
      [3, { level: 2, collapsed: false }],
      [4, { level: 2, collapsed: false }],
      [5, { level: 0, collapsed: false }],
    ])
    const state = {
      file: { sheets: [{ id: 'sh', rowCount: 10, columnCount: 10 }] },
      editJournal: { structuralOps: new Map() },
      outline: new Map([['sh', { rows: new Map(), cols }]]),
    } as unknown as LazyWorkbookState
    const { runtime } = stubRuntime(new Map())
    const lazyWorkbookRef = { current: state }
    const handle = installOutlineGutter(runtime as never, lazyWorkbookRef, {
      placement: () => ({ summaryBelow: true, summaryRight: true }),
      onToggle: vi.fn(),
    })
    expect(handle).not.toBeNull()
    await flushed(handle)
    const buttons = [...document.querySelectorAll('.outline-gutter-button')] as HTMLButtonElement[]
    // level-1 run + the level-2 run's two nested levels
    expect(buttons.length).toBe(3)
    for (const button of buttons) {
      // never above the strip's top edge onto the formula bar
      expect(Number.parseFloat(button.style.top)).toBeGreaterThanOrEqual(20)
    }
    // level 1 sits in its own lane (surface.y - 3 - 14 = 33)…
    const tops = buttons.map((button) => button.style.top).sort()
    expect(tops).toEqual(['20px', '33px', '33px'])
    // …and the level-2 button (aria-label carries its level) is the clamped one
    const deepest = buttons.find((button) =>
      button.getAttribute('aria-label')?.includes(t('appOutlineGutterCollapseCols', { n: 2 })),
    )
    expect(deepest?.style.top).toBe('20px')
    handle?.dispose()
  })
})

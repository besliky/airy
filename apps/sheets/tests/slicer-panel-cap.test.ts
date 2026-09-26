// @vitest-environment jsdom
import { beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { SlicerPanels } from '../src/renderer/SlicerPanel'
import type { TableSlicerUiState } from '../src/renderer/SlicerPanel'

/// UX-1762: a table slicer capped at TABLE_SLICER_MAX_MEMBERS must say so —
/// the panel renders an honest "+N more" line (with a hint tooltip) whenever
/// the column had more distinct values than the cap, and stays silent when it
/// did not. Pivot slicers (uncapped) never show the line.

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

function tableSlicer(moreMembers: number): TableSlicerUiState {
  return {
    id: 'tslicer-test',
    sheetId: 'main',
    tableName: 'Sales',
    colId: 0,
    fieldName: 'Key',
    members: Array.from({ length: 200 }, (_v, i) => ({ member: i, label: `v${i}` })),
    moreMembers,
    selected: Array.from({ length: 200 }, (_v, i) => i),
  }
}

describe('SlicerPanels cap indicator (UX-1762)', () => {
  it('shows +N more with the shown-of-total hint on a capped column', () => {
    const { container, unmount } = render(
      createElement(SlicerPanels, {
        slicers: [tableSlicer(9800)],
        onToggle: () => undefined,
        onSelectAll: () => undefined,
        onRemove: () => undefined,
      }),
    )
    const more = container.querySelector('.slicer-more')
    expect(more?.textContent).toBe('+9800 more')
    expect(more?.getAttribute('data-tip')).toBe(
      'Showing the first 200 of 10000 values. A selection from this list hides all values outside it.',
    )
    unmount()
  })

  it('formats the count through the locale dictionary, not a raw key', () => {
    const { container, unmount } = render(
      createElement(SlicerPanels, {
        slicers: [tableSlicer(3)],
        onToggle: () => undefined,
        onSelectAll: () => undefined,
        onRemove: () => undefined,
      }),
    )
    expect(container.querySelector('.slicer-more')?.textContent).toBe('+3 more')
    expect(container.textContent).not.toContain('dlgSlicerMore')
    unmount()
  })

  it('stays silent when the column fit under the cap', () => {
    const { container, unmount } = render(
      createElement(SlicerPanels, {
        slicers: [tableSlicer(0)],
        onToggle: () => undefined,
        onSelectAll: () => undefined,
        onRemove: () => undefined,
      }),
    )
    expect(container.querySelector('.slicer-more')).toBeNull()
    unmount()
  })
})

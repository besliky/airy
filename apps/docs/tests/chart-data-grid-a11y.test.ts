import { beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { ChartInsertModal } from '../src/renderer/components/ribbon-insert-tab'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: the en dictionary these assertions read must load first
  return loadLocale('en')
})

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

/**
 * UX-1013 (audit UX-1008): the chart dialog's data grid is legible to screen
 * readers. Every cell input carries an aria-label naming its coordinates
 * (category N / series N / "series N: <column>"), the title input is named by
 * more than its placeholder, and the native-edit disabled controls state the
 * reason in a title tooltip.
 */

describe('chart data grid accessible names (UX-1013)', () => {
  it('labels the title, header, series, and data-cell inputs', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
    })
    const container = document.createElement('div')
    document.body.appendChild(container)
    const root: Root = createRoot(container)
    act(() =>
      root.render(
        createElement(LocaleProvider, {
          initial: 'en',
          children: createElement(ChartInsertModal, { editor, onClose: vi.fn() }),
        }),
      ),
    )

    // title input: accessible name, not placeholder-only
    const inputs = [...container.querySelectorAll('input')] as HTMLInputElement[]
    const titleInput = inputs.find((i) => i.getAttribute('aria-label') === 'Chart Title')
    expect(titleInput).toBeTruthy()

    // grid: default seed is Category 1..3 columns and Series 1..2 rows
    const grid = container.querySelector('table.chart-data-grid')!
    const headerCells = [...grid.querySelectorAll('thead th input')] as HTMLInputElement[]
    expect(headerCells.map((i) => i.getAttribute('aria-label'))).toEqual([
      'Category 1',
      'Category 2',
      'Category 3',
    ])
    const firstRow = grid.querySelector('tbody tr')!
    const rowInputs = [...firstRow.querySelectorAll('input')] as HTMLInputElement[]
    expect(rowInputs[0].getAttribute('aria-label')).toBe('Series 1')
    // a data cell names its series and the column it belongs to
    expect(rowInputs[1].getAttribute('aria-label')).toBe('Series 1: Category 1')

    act(() => root.unmount())
    editor.destroy()
  })
})

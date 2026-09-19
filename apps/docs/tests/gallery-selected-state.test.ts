import { beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { SectionSettings } from '@airy-office/docx-engine'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { ColumnsDialog } from '../src/renderer/components/ColumnsDialog'
import { ChartInsertModal } from '../src/renderer/components/ribbon-insert-tab'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: the en dictionary these assertions read must load first
  return loadLocale('en')
})

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

/**
 * UX-1012 (audit UX-1004): gallery selections are announced, not just painted.
 * The selected shadow/style/chart/cols button carries aria-pressed alongside
 * its visible .active/.btn-primary pattern.
 */

function renderModal(element: React.ReactElement): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return { container, root }
}

function unmount(root: Root): void {
  act(() => root.unmount())
}

describe('gallery selected states announce via aria-pressed (UX-1012)', () => {
  it('ColumnsDialog presets mark the active column layout', () => {
    const section = {
      pageWidth: 11906,
      pageHeight: 16838,
      orientation: 'portrait',
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
      pageBorder: false,
      columns: 1,
    } as SectionSettings
    const { container, root } = renderModal(
      createElement(ColumnsDialog, { section, onApply: vi.fn(), onClose: vi.fn() }),
    )
    const presets = [...container.querySelectorAll('.cols-preset')] as HTMLButtonElement[]
    expect(presets.length).toBeGreaterThanOrEqual(5)
    // "One" is the active preset: visible class + aria-pressed
    const one = presets.find((b) => b.className.includes('active'))
    expect(one).toBe(presets[0])
    expect(one!.getAttribute('aria-pressed')).toBe('true')
    expect(presets[1].getAttribute('aria-pressed')).toBe('false')
    // picking another preset moves both the visible pattern and the aria state
    act(() => presets[1].dispatchEvent(new MouseEvent('click', { bubbles: true })))
    expect(presets[1].className).toContain('active')
    expect(presets[1].getAttribute('aria-pressed')).toBe('true')
    expect(presets[0].getAttribute('aria-pressed')).toBe('false')
    unmount(root)
  })

  it('ChartInsertModal marks the picked chart type', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
    })
    const { container, root } = renderModal(
      createElement(ChartInsertModal, { editor, onClose: vi.fn() }),
    )
    const typeRow = container.querySelector('.modal-row')!
    const kindButtons = [...typeRow.querySelectorAll('button')] as HTMLButtonElement[]
    expect(kindButtons.length).toBe(7)
    // default kind is bar: visible .btn-primary + aria-pressed, others off
    const pressed = kindButtons.filter((b) => b.getAttribute('aria-pressed') === 'true')
    expect(pressed).toHaveLength(1)
    expect(pressed[0].className).toContain('btn-primary')
    act(() => kindButtons[1].dispatchEvent(new MouseEvent('click', { bubbles: true })))
    expect(kindButtons[1].getAttribute('aria-pressed')).toBe('true')
    expect(kindButtons[0].getAttribute('aria-pressed')).toBe('false')
    unmount(root)
    editor.destroy()
  })
})

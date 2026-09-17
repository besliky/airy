// @vitest-environment jsdom
import { beforeAll, describe, expect, it } from 'vitest'
import { act, createElement, useState } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { ribbonPanelProps, useRibbonTablist } from '../src/ribbon-tablist'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

/* Mounts the smallest possible tablist the ribbons build: a wrapper div with
 * the tablist props, one button per tab with the tab props, a panel div. */
function Probe({ tabs }: { tabs: readonly string[] }) {
  const [active, setActive] = useState<string | null>(tabs[0] ?? null)
  const list = useRibbonTablist({
    tabs,
    activeTab: active,
    idPrefix: 'test-ribbon',
    label: 'Ribbon tabs',
    onSelect: setActive,
  })
  return createElement(
    'div',
    null,
    createElement(
      'div',
      { className: 'ribbon-tablist', ...list.tablistProps },
      tabs.map((tab) => createElement('button', { key: tab, ...list.tabProps(tab) }, tab)),
    ),
    createElement('div', ribbonPanelProps('test-ribbon', active)),
  )
}

function mount(tabs: readonly string[]): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(Probe, { tabs })))
  return { root, container }
}

function key(element: Element, keyName: string): void {
  act(() => {
    element.dispatchEvent(new KeyboardEvent('keydown', { key: keyName, bubbles: true }))
  })
}

describe('useRibbonTablist', () => {
  it('exposes the tablist ARIA structure with a roving tabindex', () => {
    const { root, container } = mount(['home', 'insert', 'draw'])
    const strip = container.querySelector('.ribbon-tablist') as HTMLElement
    expect(strip.getAttribute('role')).toBe('tablist')
    expect(strip.getAttribute('aria-label')).toBe('Ribbon tabs')

    const buttons = Array.from(container.querySelectorAll('button'))
    expect(buttons.map((b) => b.getAttribute('role'))).toEqual(['tab', 'tab', 'tab'])
    expect(buttons.map((b) => b.tabIndex)).toEqual([0, -1, -1])
    expect(buttons.map((b) => b.getAttribute('aria-selected'))).toEqual(['true', 'false', 'false'])
    expect(buttons[0].id).toBe('test-ribbon-tab-home')

    const panel = container.querySelector('[role="tabpanel"]') as HTMLElement
    expect(panel.id).toBe('test-ribbon-panel')
    expect(panel.getAttribute('aria-labelledby')).toBe('test-ribbon-tab-home')
    act(() => root.unmount())
  })

  it('arrow keys switch the tab and move focus (automatic activation)', () => {
    const { root, container } = mount(['home', 'insert', 'draw'])
    const buttons = () => Array.from(container.querySelectorAll('button'))
    buttons()[0].focus()
    expect(document.activeElement).toBe(buttons()[0])

    key(buttons()[0], 'ArrowRight')
    expect(buttons().map((b) => b.getAttribute('aria-selected'))).toEqual([
      'false',
      'true',
      'false',
    ])
    expect(document.activeElement).toBe(buttons()[1])
    // roving tabindex follows the selection
    expect(buttons().map((b) => b.tabIndex)).toEqual([-1, 0, -1])

    key(buttons()[1], 'ArrowRight')
    key(buttons()[2], 'ArrowRight')
    // wrapped back to the first tab
    expect(buttons().map((b) => b.getAttribute('aria-selected'))).toEqual([
      'true',
      'false',
      'false',
    ])
    expect(document.activeElement).toBe(buttons()[0])

    key(buttons()[0], 'ArrowLeft')
    expect(document.activeElement).toBe(buttons()[2])
    act(() => root.unmount())
  })

  it('Home and End jump to the strip ends', () => {
    const { root, container } = mount(['home', 'insert', 'draw'])
    const buttons = () => Array.from(container.querySelectorAll('button'))
    buttons()[0].focus()
    key(buttons()[0], 'End')
    expect(document.activeElement).toBe(buttons()[2])
    key(buttons()[2], 'Home')
    expect(document.activeElement).toBe(buttons()[0])
    act(() => root.unmount())
  })

  it('leaves keys the tablist does not own untouched', () => {
    const { root, container } = mount(['home', 'insert', 'draw'])
    const first = container.querySelectorAll('button')[0]
    first.focus()
    let defaultNotPrevented = false
    first.addEventListener('keydown', (e) => {
      defaultNotPrevented = !e.defaultPrevented
    })
    key(first, 'ArrowDown')
    expect(defaultNotPrevented).toBe(true)
    expect(first.getAttribute('aria-selected')).toBe('true')
    act(() => root.unmount())
  })
})

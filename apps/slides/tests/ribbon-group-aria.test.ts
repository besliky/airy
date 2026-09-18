// Slides ribbon groups are named ARIA groups: the shared Group component
// renders role=group + aria-label (the localized visible label) in both the
// expanded and the collapsed dropdown form, mirroring the sheets ribbon
// groups' named <section> wrappers.
import { beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { Group } from '../src/renderer/components/ribbon-shared'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

function mount(element: React.ReactElement): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return { container, root }
}

function unmount(container: HTMLElement, root: Root) {
  act(() => root.unmount())
  container.remove()
}

describe('ribbon Group aria structure', () => {
  it('names the expanded group by its visible label', () => {
    const { container, root } = mount(
      createElement(Group, {
        label: 'Clipboard',
        groupId: 'clip',
        children: createElement('button', null, 'Copy'),
      }),
    )
    const group = container.querySelector('.ribbon-group') as HTMLElement
    expect(group.getAttribute('role')).toBe('group')
    expect(group.getAttribute('aria-label')).toBe('Clipboard')
    expect(group.getAttribute('data-rbgroup')).toBe('clip')
    // the visible label stays rendered for sighted users
    expect(group.querySelector('.ribbon-group-label')?.textContent).toBe('Clipboard')
    unmount(container, root)
  })

  it('keeps the naming when the group collapses into its dropdown form', () => {
    const { container, root } = mount(
      createElement(Group, {
        label: 'Font',
        groupId: 'font',
        collapse: { collapsed: true, open: false, onToggle: () => {}, icon: null },
        children: createElement('button', null, 'Bold'),
      }),
    )
    const group = container.querySelector('.ribbon-group') as HTMLElement
    expect(group.getAttribute('role')).toBe('group')
    expect(group.getAttribute('aria-label')).toBe('Font')
    // collapsed groups surface their contents through the toggle button
    expect(group.querySelector('button.rb-big')).not.toBeNull()
    unmount(container, root)
  })
})

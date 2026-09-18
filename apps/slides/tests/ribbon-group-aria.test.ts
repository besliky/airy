// Slides ribbon groups are named ARIA groups: the shared Group component
// names itself with aria-labelledby pointing at its visible label element in
// both the expanded and the collapsed dropdown form (a single source of
// truth — a duplicated aria-label could drift from the visible text),
// mirroring the sheets ribbon groups' named <section> wrappers.
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
  it('names the expanded group through its visible label element', () => {
    const { container, root } = mount(
      createElement(Group, {
        label: 'Clipboard',
        groupId: 'clip',
        children: createElement('button', null, 'Copy'),
      }),
    )
    const group = container.querySelector('.ribbon-group') as HTMLElement
    expect(group.getAttribute('role')).toBe('group')
    // the name comes from the label element, not from a second copy of the text
    expect(group.getAttribute('aria-label')).toBeNull()
    const labelId = group.getAttribute('aria-labelledby')
    expect(labelId).toBeTruthy()
    const label = group.querySelector('.ribbon-group-label') as HTMLElement
    expect(label.getAttribute('id')).toBe(labelId)
    // the accessible name resolves to the visible text
    expect(label.textContent).toBe('Clipboard')
    expect(group.getAttribute('data-rbgroup')).toBe('clip')
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
    const labelId = group.getAttribute('aria-labelledby')
    expect(labelId).toBeTruthy()
    expect((group.querySelector('.ribbon-group-label') as HTMLElement).getAttribute('id')).toBe(
      labelId,
    )
    // collapsed groups surface their contents through the toggle button
    expect(group.querySelector('button.rb-big')).not.toBeNull()
    unmount(container, root)
  })

  it('gives sibling groups distinct label ids', () => {
    const { container, root } = mount(
      createElement(
        'div',
        null,
        createElement(Group, {
          label: 'Clipboard',
          groupId: 'clip',
          children: createElement('button', null, 'Copy'),
        }),
        createElement(Group, {
          label: 'Clipboard',
          groupId: 'clip2',
          children: createElement('button', null, 'Paste'),
        }),
      ),
    )
    const ids = [...container.querySelectorAll('.ribbon-group-label')].map((label) =>
      label.getAttribute('id'),
    )
    expect(ids[0]).toBeTruthy()
    expect(ids[1]).toBeTruthy()
    expect(ids[0]).not.toBe(ids[1])
    unmount(container, root)
  })
})

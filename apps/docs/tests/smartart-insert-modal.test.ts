// TEST-702: the SmartArt insert gallery UI (ribbon-insert-tab) had no tests —
// the engine below it was covered, the dialog was not. These DOM tests pin the
// behaviors Word users touch directly: gallery preset selection (a functional
// radio group), per-preset node defaults, the level clamps in the hierarchy
// editor, the live preview miniatures, the 8-node UI cap and the insert path
// that lands a diagram block in the editor.
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { LocaleProvider, loadLocale, setModuleLang } from '../src/renderer/i18n/locale'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { SmartArtInsertModal } from '../src/renderer/components/ribbon-insert-tab'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
  // PERF-904: the en dictionary these assertions read must load first
  return loadLocale('en')
})

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

function openModal(): {
  container: HTMLElement
  root: Root
  editor: Editor
  onClose: ReturnType<typeof vi.fn>
} {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
  })
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  const onClose = vi.fn()
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(SmartArtInsertModal, { editor, onClose }),
      }),
    ),
  )
  return { container, root, editor, onClose }
}

/** React-controlled input: set the value through the native setter so onChange fires */
function typeText(input: HTMLInputElement, value: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  act(() => {
    setter.call(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function nodeInputs(container: HTMLElement): HTMLInputElement[] {
  return [...container.querySelectorAll('.smartart-node-row input')] as HTMLInputElement[]
}

function cellButtons(container: HTMLElement): HTMLButtonElement[] {
  return [
    ...container.querySelectorAll('.smartart-gallery button[role="radio"]'),
  ] as HTMLButtonElement[]
}

describe('SmartArt insert gallery (TEST-702)', () => {
  it('shows the four presets as a radio group with thumbnails, block list preselected', () => {
    const { container, root, editor } = openModal()
    expect(container.querySelector('.smartart-gallery')!.getAttribute('role')).toBe('radiogroup')
    const cells = cellButtons(container)
    expect(cells.map((c) => c.querySelector('.smartart-cell-label')!.textContent)).toEqual([
      'Basic Block List',
      'Vertical Bullet List',
      'Basic Process',
      'Organization Chart',
    ])
    expect(cells.map((c) => c.getAttribute('aria-checked'))).toEqual([
      'true',
      'false',
      'false',
      'false',
    ])
    expect(cells[0]!.className).toContain('selected')
    // every cell previews its layout through the same shape model as the editor
    for (const cell of cells) {
      expect(cell.querySelectorAll('.smartart-thumb-shape').length).toBeGreaterThan(0)
    }
    act(() => root.unmount())
    editor.destroy()
  })

  it("switching presets resets the node list to that layout's defaults", () => {
    const { container, root, editor } = openModal()
    const hier = cellButtons(container)[3]!
    act(() => hier.click())
    expect(hier.getAttribute('aria-checked')).toBe('true')
    expect(hier.className).toContain('selected')
    const inputs = nodeInputs(container)
    // the classic 5-node org tree: levels 0,1,2,1,1 rendered as indents
    expect(inputs.map((i) => i.value)).toEqual(['Item 1', 'Item 2', 'Item 3', 'Item 4', 'Item 5'])
    expect(
      [...container.querySelectorAll('.smartart-node-row')].map(
        (row) => (row as HTMLElement).style.paddingLeft,
      ),
    ).toEqual(['0px', '22px', '44px', '22px', '22px'])
    // back to a flat preset: three level-0 nodes again
    act(() => cellButtons(container)[0]!.click())
    expect(nodeInputs(container).map((i) => i.value)).toEqual(['Item 1', 'Item 2', 'Item 3'])
    expect(
      nodeInputs(container).every(
        (i) => (i.closest('.smartart-node-row') as HTMLElement).style.paddingLeft === '0px',
      ),
    ).toBe(true)
    act(() => root.unmount())
    editor.destroy()
  })

  it('clamps hierarchy levels: no jump deeper than one, outdent stops at 0', () => {
    const { container, root, editor } = openModal()
    act(() => cellButtons(container)[3]!.click())
    const levelButton = (row: number, label: 'Indent' | 'Outdent') =>
      [...container.querySelectorAll('.smartart-node-row')[row]!.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === label,
      ) as HTMLButtonElement

    // row 0 can never indent (first node is level 0); row 1 sits at the
    // maximum depth already (previous level 0 + 1)
    expect(levelButton(0, 'Indent').disabled).toBe(true)
    expect(levelButton(1, 'Indent').disabled).toBe(true)
    // row 3 (level 1 under a level-2 parent) may go one deeper
    expect(levelButton(3, 'Indent').disabled).toBe(false)

    // outdenting node 3 (index 2, level 2) walks down to level 0 and stops:
    // the button disables at the floor instead of producing a negative level
    act(() => levelButton(2, 'Outdent').click())
    expect(levelButton(2, 'Outdent').disabled).toBe(false)
    act(() => levelButton(2, 'Outdent').click())
    expect(levelButton(2, 'Outdent').disabled).toBe(true)
    expect(
      ([...container.querySelectorAll('.smartart-node-row')][2] as HTMLElement).style.paddingLeft,
    ).toBe('0px')
    act(() => root.unmount())
    editor.destroy()
  })

  it('previews the live node model: texts follow edits, hierarchy adds connectors', () => {
    const { container, root, editor } = openModal()
    const preview = () => container.querySelector('.smartart-preview .smartart-thumb')!
    // flat block list: shapes only, no connector rules
    expect(preview().querySelectorAll('.smartart-thumb-shape').length).toBe(3)
    expect(preview().querySelectorAll('.smartart-thumb-rule').length).toBe(0)
    expect(
      [...preview().querySelectorAll('.smartart-thumb-text')].map((t) => t.textContent),
    ).toEqual(['Item 1', 'Item 2', 'Item 3'])

    // typing re-keys the preview through buildDiagramDisplay
    typeText(nodeInputs(container)[0]!, 'Alpha')
    expect(
      [...preview().querySelectorAll('.smartart-thumb-text')].map((t) => t.textContent),
    ).toEqual(['Alpha', 'Item 2', 'Item 3'])

    // the org chart layout draws elbow connectors between levels
    act(() => cellButtons(container)[3]!.click())
    expect(preview().querySelectorAll('.smartart-thumb-rule').length).toBeGreaterThan(0)
    act(() => root.unmount())
    editor.destroy()
  })

  it('caps the node list at 8 rows and disables the add button there', () => {
    const { container, root, editor } = openModal()
    const add = [...container.querySelectorAll('.modal-row button')].find(
      (b) => b.textContent === '+ Node',
    ) as HTMLButtonElement
    for (let i = 0; i < 5; i++) act(() => add.click())
    expect(nodeInputs(container).length).toBe(8)
    expect(add.disabled).toBe(true)
    act(() => add.click())
    expect(nodeInputs(container).length).toBe(8)
    // and the last row cannot be removed while one node remains
    const remove = (last: Element) =>
      [...last.querySelectorAll('button')].find(
        (b) => b.getAttribute('aria-label') === 'Remove node',
      ) as HTMLButtonElement
    for (let i = 0; i < 7; i++)
      act(() => remove(container.querySelector('.smartart-node-row:last-child')!).click())
    expect(nodeInputs(container).length).toBe(1)
    expect(remove(container.querySelector('.smartart-node-row:last-child')!).disabled).toBe(true)
    act(() => root.unmount())
    editor.destroy()
  })

  it('inserts a diagram block from the live items and closes; refuses an all-empty model', () => {
    const { container, root, editor, onClose } = openModal()
    typeText(nodeInputs(container)[0]!, 'Alpha')
    act(() => {
      ;(container.querySelector('.modal-actions .btn-primary') as HTMLButtonElement).click()
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    const inserted = (editor.getJSON() as { content: Array<Record<string, unknown>> }).content.find(
      (node) => node.type === 'docProtected',
    )
    expect(inserted).toBeTruthy()
    expect(inserted!.attrs).toMatchObject({
      blockType: 'diagram',
      genDiagram: {
        kind: 'blockList',
        items: [
          { text: 'Alpha', level: 0 },
          { text: 'Item 2', level: 0 },
          { text: 'Item 3', level: 0 },
        ],
      },
    })
    act(() => root.unmount())
    editor.destroy()

    // clearing every node text leaves nothing to insert: the primary button
    // disables instead of writing an empty diagram
    const second = openModal()
    for (const input of nodeInputs(second.container)) typeText(input, '  ')
    expect(
      (second.container.querySelector('.modal-actions .btn-primary') as HTMLButtonElement).disabled,
    ).toBe(true)
    act(() => {
      ;(second.container.querySelector('.modal-actions .btn-primary') as HTMLButtonElement).click()
    })
    expect(second.onClose).not.toHaveBeenCalled()
    act(() => second.root.unmount())
    second.editor.destroy()
  })
})

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { FindPanel, findMatches, foldCase } from '../src/renderer/components/FindPanel'
import { clampPictureCm, PICTURE_CM_MAX, PICTURE_CM_MIN } from '../src/renderer/components/Ribbon'

function createEditor(text: string): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text }],
        },
      ],
    },
  })
}

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

beforeEach(() => {
  Element.prototype.scrollIntoView = vi.fn()
})

describe('foldCase', () => {
  it('lowercases without changing string length', () => {
    expect(foldCase('ABC def')).toBe('abc def')
    // 'İ'.toLowerCase() is 2 chars ('i' + combining dot) and would shift offsets
    expect('İ'.toLowerCase().length).toBe(2)
    expect(foldCase('İstanbul').length).toBe('İstanbul'.length)
    expect(foldCase('AİB')).toBe('aİb')
  })
})

describe('findMatches', () => {
  it('keeps offsets aligned after a length-changing lowercase char', () => {
    const editor = createEditor('İİİ test')
    const [m] = findMatches(editor, 'TEST', {
      matchCase: false,
      wholeWord: false,
      useWildcards: false,
      ignoreDiacritics: false,
    })
    expect(m).toBeDefined()
    expect(editor.state.doc.textBetween(m.from, m.to)).toBe('test')
    editor.destroy()
  })

  it('respects matchCase and wholeWord', () => {
    const editor = createEditor('Cat cats CAT')
    expect(
      findMatches(editor, 'cat', {
        matchCase: false,
        wholeWord: false,
        useWildcards: false,
        ignoreDiacritics: false,
      }),
    ).toHaveLength(3)
    expect(
      findMatches(editor, 'cat', {
        matchCase: false,
        wholeWord: true,
        useWildcards: false,
        ignoreDiacritics: false,
      }),
    ).toHaveLength(2)
    expect(
      findMatches(editor, 'CAT', {
        matchCase: true,
        wholeWord: false,
        useWildcards: false,
        ignoreDiacritics: false,
      }),
    ).toHaveLength(1)
    editor.destroy()
  })
})

describe('clampPictureCm', () => {
  it('clamps to the Word picture size range', () => {
    expect(clampPictureCm(0.001)).toBe(PICTURE_CM_MIN)
    expect(clampPictureCm(999)).toBe(PICTURE_CM_MAX)
    expect(clampPictureCm(10)).toBe(10)
  })
})

describe('FindPanel', () => {
  afterEach(() => {
    vi.useRealTimers()
  })

  it('hides the replace row when the editor is read-only', () => {
    const editor = createEditor('hello world')
    editor.setEditable(false)
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    expect(container.querySelectorAll('.find-row')).toHaveLength(1)
    expect(container.querySelector('.find-action')).toBeNull()
    unmount()
    editor.destroy()
  })

  it('shows the replace row on an editable document', () => {
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    expect(container.querySelectorAll('.find-row')).toHaveLength(2)
    unmount()
    editor.destroy()
  })

  it('debounces the scan while typing', () => {
    vi.useFakeTimers()
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    const input = container.querySelector<HTMLInputElement>('.find-input')!
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(input, 'hello')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const count = () => container.querySelector('.find-count')!.textContent
    expect(count()).not.toContain('1/1')
    act(() => {
      vi.advanceTimersByTime(200)
    })
    expect(count()).toBe('1/1')
    unmount()
    editor.destroy()
  })

  const optButton = (container: HTMLElement, label: string) =>
    [...container.querySelectorAll<HTMLButtonElement>('.find-opt')].find(
      (b) => b.textContent === label,
    )!

  const click = (el: Element) =>
    act(() => {
      el.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

  it('disables the whole-word toggle while wildcards are on', () => {
    const editor = createEditor('hello world')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    const word = optButton(container, 'W')
    const wild = optButton(container, '*?')
    expect(word.disabled).toBe(false)
    click(wild)
    expect(word.disabled).toBe(true)
    expect(wild.className).toContain('on')
    click(wild)
    expect(word.disabled).toBe(false)
    unmount()
    editor.destroy()
  })

  it('finds and replaces using wildcards', () => {
    vi.useFakeTimers()
    const editor = createEditor('hello hallo')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    click(optButton(container, '*?'))
    const input = container.querySelector<HTMLInputElement>('.find-input')!
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(input, 'h?llo')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    act(() => {
      vi.advanceTimersByTime(200)
    })
    expect(container.querySelector('.find-count')!.textContent).toBe('1/2')
    const replaceInput = container.querySelectorAll<HTMLInputElement>('.find-input')[1]!
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(replaceInput, 'X')
      replaceInput.dispatchEvent(new Event('input', { bubbles: true }))
    })
    click(container.querySelectorAll('.find-action')[1]!) // Replace All
    expect(editor.state.doc.textContent).toBe('X X')
    unmount()
    editor.destroy()
  })

  it('ignore-diacritics widens the count and yields to match case', () => {
    vi.useFakeTimers()
    const editor = createEditor('café cafe')
    const { container, unmount } = render(createElement(FindPanel, { editor, onClose: () => {} }))
    const input = container.querySelector<HTMLInputElement>('.find-input')!
    const type = (el: HTMLInputElement, value: string) =>
      act(() => {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
        setter.call(el, value)
        el.dispatchEvent(new Event('input', { bubbles: true }))
      })
    type(input, 'cafe')
    act(() => {
      vi.advanceTimersByTime(200)
    })
    expect(container.querySelector('.find-count')!.textContent).toBe('1/1') // café differs
    const dia = optButton(container, 'é')
    click(dia)
    expect(container.querySelector('.find-count')!.textContent).toBe('1/2')
    const caseBtn = optButton(container, 'Aa')
    click(caseBtn)
    expect(dia.disabled).toBe(true) // accents always differ when match case is on
    expect(container.querySelector('.find-count')!.textContent).toBe('1/1')
    unmount()
    editor.destroy()
  })
})

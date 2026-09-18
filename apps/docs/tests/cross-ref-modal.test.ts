/**
 * CrossRefModal DOM contract: dialog semantics (UX-902) — the cross-reference
 * dialog carries full modal semantics via the suite's useModalDialog hook:
 * role="dialog" / aria-modal named by its visible heading, focus lands inside
 * on open, Tab wraps at the dialog edges, Escape closes it from anywhere
 * inside. Insert flow (BUG-918) — the insert is gated on editability and
 * lands the hidden-anchor stamp together with the REF in one transaction, so
 * one undo removes both.
 */
import { Editor } from '@tiptap/core'
import type { Block } from '@airy-office/docx-engine'
import { parseDocx } from '@airy-office/docx-engine'
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { CrossRefModal } from '../src/renderer/components/ribbon-insert-tab'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter One</w:t></w:r></w:p>' +
  '<w:p><w:bookmarkStart w:id="1" w:name="Intro"/><w:bookmarkEnd w:id="1"/>' +
  '<w:r><w:t>Intro paragraph.</w:t></w:r></w:p>'

const editors = new Set<Editor>()

async function openDoc(bodyXml = BODY) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, blocks: parsed.blocks }
}

describe('CrossRefModal dialog semantics (UX-902)', () => {
  let editor: Editor
  let blocks: Block[]
  let container: HTMLElement
  let root: Root

  beforeEach(async () => {
    ;({ editor, blocks } = await openDoc())
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
  })

  const mountModal = (onClose: () => void) =>
    act(() => root.render(createElement(CrossRefModal, { editor, blocks, onClose })))

  it('is a role=dialog named by its heading with aria-modal', () => {
    mountModal(vi.fn())
    const modal = container.querySelector<HTMLElement>('.modal')!
    const title = modal.querySelector<HTMLElement>('h2')!
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe(t('ribbonCrossRef'))
  })

  it('moves focus into the dialog on open and closes on Escape', () => {
    const onClose = vi.fn()
    mountModal(onClose)
    const backdrop = container.querySelector<HTMLElement>('.modal-backdrop')!
    expect(backdrop.contains(document.activeElement)).toBe(true)
    act(() => {
      backdrop
        .querySelector('h2')!
        .dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('wraps Tab focus at the dialog edges (focus trap)', () => {
    mountModal(vi.fn())
    const backdrop = container.querySelector<HTMLElement>('.modal-backdrop')!
    const buttons = [...backdrop.querySelectorAll<HTMLButtonElement>('button:not([disabled])')]
    const last = buttons[buttons.length - 1]!
    last.focus()
    expect(document.activeElement).toBe(last)
    act(() => {
      last.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }))
    })
    // Tab from the last control wraps around to the first one instead of
    // escaping to the page behind the backdrop
    expect(document.activeElement).not.toBe(last)
    expect(backdrop.contains(document.activeElement)).toBe(true)
  })
})

describe('CrossRefModal insert flow (BUG-918)', () => {
  let editor: Editor
  let blocks: Block[]
  let container: HTMLElement
  let root: Root
  let onClose: Mock<() => void>

  beforeEach(async () => {
    ;({ editor, blocks } = await openDoc())
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    onClose = vi.fn<() => void>()
    act(() => root.render(createElement(CrossRefModal, { editor, blocks, onClose })))
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
  })

  const headingRow = () =>
    [...container.querySelectorAll<HTMLButtonElement>('.bookmark-list .bookmark-name')].find(
      (b) => b.textContent === 'Chapter One',
    )!

  const refFieldTexts = () => {
    const texts: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.marks.some((m) => m.type.name === 'refField')) texts.push(node.textContent)
    })
    return texts
  }

  it('inserts a heading reference and its hidden anchor as ONE undo step', () => {
    const before = editor.state.doc.toJSON()
    let txCount = 0
    editor.on('transaction', ({ transaction }) => {
      if (transaction.docChanged) txCount += 1
    })
    act(() => headingRow().click())
    expect(onClose).toHaveBeenCalledTimes(1)
    // the anchor stamp and the REF insert coalesce into a single mutating
    // transaction — one history item by construction, independent of the
    // history plugin's time/selection grouping rules (the trailing
    // focus-only transaction carries no steps and no history item)
    expect(txCount).toBe(1)
    // the REF field landed with the heading text as its cache…
    expect(refFieldTexts()).toEqual(['Chapter One'])
    // …and the heading carries the stamped `_Toc…` anchor it points at
    const heading = editor.state.doc.firstChild!
    const hidden = (heading.attrs.hiddenBookmarks as string[] | null) ?? []
    const anchor = hidden.find((n) => /^_Toc\d+$/.test(n))
    expect(anchor).toMatch(/^_Toc\d{9}$/)
    const refNames: string[] = []
    editor.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        if (mark.type.name === 'refField') refNames.push(String(mark.attrs.name))
      }
    })
    expect(refNames).toEqual([anchor])
    // a single Ctrl+Z removes reference AND anchor together — no stray
    // hidden bookmark left behind to save into the file
    act(() => {
      editor.commands.undo()
    })
    expect(refFieldTexts()).toEqual([])
    expect(editor.state.doc.toJSON()).toEqual(before)
  })

  it('is a complete no-op on a read-only editor (UX-715 gate)', () => {
    editor.setEditable(false)
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => undefined)
    const docBefore = editor.state.doc
    act(() => headingRow().click())
    // the guard runs before any anchor work: no transaction, no insert, no
    // misleading alert — and the dialog stays open instead of silently
    // pretending success (the bookmark-kind path used to insert and close)
    expect(editor.state.doc).toBe(docBefore)
    expect(refFieldTexts()).toEqual([])
    expect(onClose).not.toHaveBeenCalled()
    expect(alert).not.toHaveBeenCalled()
    expect(editor.state.doc.firstChild!.attrs.hiddenBookmarks).toBeNull()
    alert.mockRestore()
  })
})

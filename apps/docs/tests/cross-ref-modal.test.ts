/**
 * CrossRefModal DOM contract (UX-902): the cross-reference dialog carries full
 * modal semantics via the suite's useModalDialog hook — role="dialog" /
 * aria-modal named by its visible heading, focus lands inside on open, Tab
 * wraps at the dialog edges, and Escape closes it from anywhere inside.
 */
import { Editor } from '@tiptap/core'
import type { Block } from '@airy-office/docx-engine'
import { parseDocx } from '@airy-office/docx-engine'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
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

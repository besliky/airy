import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { ribbonEscapeDeferred } from '../src/renderer/components/ribbon-shared'

/**
 * UX-11s2: slides ribbon popups closed on outside press / window blur / the
 * chrome-pressed relay, but not on Escape — a mouse-only dismissal, while a
 * keyboard user who Tabbed into a trigger had no way back out. Ribbon.tsx now
 * installs a window-capture Escape handler that claims the key (the global
 * Esc actions respect defaultPrevented) and collapses every panel. This file
 * pins the deferral predicate — the cases where a layer beneath the popup
 * owns the press — and the wiring, without mounting the whole Ribbon.
 */

/** mount `html` inside a `.rb-drop-wrap` ribbon widget and return the probe element */
function inWidget(html: string): HTMLElement {
  const wrap = document.createElement('div')
  wrap.className = 'rb-drop-wrap'
  wrap.innerHTML = html
  document.body.appendChild(wrap)
  return wrap.querySelector('[data-probe]')!
}

describe('ribbonEscapeDeferred (UX-11s2)', () => {
  it('claims the press from a plain trigger button', () => {
    const btn = inWidget('<button data-probe class="rb-big rb-split"><span>Split</span></button>')
    expect(ribbonEscapeDeferred(btn)).toBe(false)
  })

  it('claims the press from a panel item (a button inside the open popup)', () => {
    const item = inWidget(
      '<div class="rb-drop rb-menu"><button data-probe>From Beginning</button></div>',
    )
    expect(ribbonEscapeDeferred(item)).toBe(false)
  })

  it('defers to a modal dialog stacked on top of the ribbon', () => {
    const backdrop = document.createElement('div')
    backdrop.className = 'modal-backdrop'
    const dlg = document.createElement('div')
    dlg.className = 'modal'
    const input = document.createElement('input')
    dlg.appendChild(input)
    backdrop.appendChild(dlg)
    document.body.appendChild(backdrop)
    expect(ribbonEscapeDeferred(input)).toBe(true)
    expect(ribbonEscapeDeferred(dlg)).toBe(true)
    document.body.innerHTML = ''
  })

  it('defers to an editable ribbon field cancelling its draft (font combobox)', () => {
    const input = inWidget(
      '<div class="rb-font-btn"><input data-probe class="rb-font-input" /></div>',
    )
    expect(ribbonEscapeDeferred(input)).toBe(true)
  })

  it('claims the press from a non-editable inside the widget (size caret button)', () => {
    const caret = inWidget('<button data-probe class="rb-size-caret" aria-expanded="true" />')
    expect(ribbonEscapeDeferred(caret)).toBe(false)
  })

  it('defers to an expanded shared Dropdown list (it closes itself first)', () => {
    const item = inWidget(
      '<span class="gs-dd"><button aria-expanded="true" data-probe /><div class="gs-dd-pop" /></span>',
    )
    expect(ribbonEscapeDeferred(item)).toBe(true)
  })

  it('claims the press from a collapsed shared Dropdown trigger', () => {
    const item = inWidget('<span class="gs-dd"><button aria-expanded="false" data-probe /></span>')
    expect(ribbonEscapeDeferred(item)).toBe(false)
  })

  it('tolerates a null target (focus nowhere)', () => {
    expect(ribbonEscapeDeferred(null)).toBe(false)
  })
})

describe('Ribbon wires the escape dismissal (source contract)', () => {
  const src = readFileSync(join(__dirname, '../src/renderer/components/Ribbon.tsx'), 'utf8')

  it('installs the capture-phase handler while any panel is open', () => {
    expect(src).toContain('ribbonEscapeDeferred(e.target as Element | null)')
    expect(src).toMatch(/addEventListener\('keydown', onKey, true\)/)
  })

  it('claims the key so the global Esc actions stand down', () => {
    const handler = src.slice(
      src.indexOf('const onKey = (e: KeyboardEvent) => {', src.indexOf('ribbonEscapeDeferred')),
      src.indexOf('closePanels()', src.indexOf('ribbonEscapeDeferred')),
    )
    expect(handler).toContain('e.preventDefault()')
    expect(handler).toContain('e.stopPropagation()')
  })
})

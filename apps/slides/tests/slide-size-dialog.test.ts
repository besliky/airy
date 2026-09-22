import { beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { LocaleProvider, setModuleLang, t } from '../src/renderer/i18n/locale'
import {
  SlideSizeDialog,
  clampLengthCm,
  MAX_SLIDE_CM,
  MIN_SLIDE_CM,
  toEmu,
} from '../src/renderer/components/SlideSizeDialog'

/**
 * PAR-316 custom slide size: the Design → Slide Size dropdown gains a
 * "Custom…" dialog (PowerPoint parity). Covers the unit/EMU conversion, the
 * PowerPoint bounds clamp, the modal contract basics and the OK flow
 * (landscape/portrait swap → EMU, then onClose).
 */
beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { slidesApi: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

/** Simulate typing into React's controlled number input */
function typeInto(input: HTMLInputElement, text: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  act(() => {
    setter.call(input, text)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

describe('slide size unit conversion', () => {
  it('converts cm and inches to EMU', () => {
    expect(toEmu(1, 'cm')).toBe(360000)
    expect(toEmu(1, 'in')).toBe(914400)
    expect(toEmu(25.4, 'cm')).toBe(9144000)
  })

  it('clamps into the PowerPoint bounds (0.5in..56in)', () => {
    expect(MIN_SLIDE_CM).toBeCloseTo(1.27)
    expect(MAX_SLIDE_CM).toBeCloseTo(142.24)
    expect(clampLengthCm(0)).toBe(MIN_SLIDE_CM)
    expect(clampLengthCm(999)).toBe(MAX_SLIDE_CM)
    expect(clampLengthCm(25.4)).toBe(25.4)
  })
})

describe('SlideSizeDialog', () => {
  const renderDialog = (onApply = vi.fn(), onClose = vi.fn()) =>
    render(
      createElement(SlideSizeDialog, {
        widthCm: 33.87,
        heightCm: 19.05,
        onApply,
        onClose,
      }),
    )

  it('renders as a dialog named by its heading', () => {
    const { container, unmount } = renderDialog()
    const modal = container.querySelector('.modal') as HTMLElement
    const title = container.querySelector('.modal h2') as HTMLElement
    expect(modal.getAttribute('role')).toBe('dialog')
    expect(modal.getAttribute('aria-modal')).toBe('true')
    expect(modal.getAttribute('aria-labelledby')).toBe(title.id)
    expect(title.textContent).toBe(t('ribbonSlideSize'))
    unmount()
  })

  it('applies landscape EMU from cm values and closes', () => {
    const onApply = vi.fn()
    const onClose = vi.fn()
    const { container, unmount } = renderDialog(onApply, onClose)
    const buttons = [...container.querySelectorAll<HTMLButtonElement>('.modal-actions button')]
    act(() => buttons[1]!.click())
    // 33.87cm x 19.05cm ≈ the 16:9 12192000x6858000 EMU slide
    expect(onApply).toHaveBeenCalledWith(12193200, 6858000)
    expect(onClose).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('swaps width/height for portrait', () => {
    const onApply = vi.fn()
    const { container, unmount } = renderDialog(onApply, vi.fn())
    const portrait = container.querySelector<HTMLInputElement>(
      'input[name="slide-size-orientation"]:nth-of-type(1)',
    )
    const radios = [
      ...container.querySelectorAll<HTMLInputElement>('input[name="slide-size-orientation"]'),
    ]
    act(() => {
      ;(portrait ?? radios[1]).click()
      radios[1]!.click()
    })
    const ok = [...container.querySelectorAll<HTMLButtonElement>('.modal-actions button')][1]!
    act(() => ok.click())
    expect(onApply).toHaveBeenCalledWith(6858000, 12193200)
    unmount()
  })

  it('disables OK and shows the range error outside the PowerPoint bounds', () => {
    const { container, unmount } = renderDialog()
    const inputs = [...container.querySelectorAll<HTMLInputElement>('.slide-size-fields input')]
    typeInto(inputs[0]!, '200')
    const ok = [...container.querySelectorAll<HTMLButtonElement>('.modal-actions button')][1]!
    expect(ok.disabled).toBe(true)
    expect(container.textContent).toContain(t('appSlideSizeRange'))
    // back inside the bounds → OK enabled again
    typeInto(inputs[0]!, '25')
    expect(ok.disabled).toBe(false)
    unmount()
  })

  it('converts entered values when switching between cm and in', () => {
    const { container, unmount } = renderDialog()
    const inputs = [...container.querySelectorAll<HTMLInputElement>('.slide-size-fields input')]
    const unitButtons = [
      ...container.querySelectorAll<HTMLButtonElement>('.slide-size-unit button'),
    ]
    act(() => unitButtons[1]!.click()) // cm → in
    expect(Number(inputs[0]!.value)).toBeCloseTo(33.87 / 2.54, 1)
    expect(Number(inputs[1]!.value)).toBeCloseTo(19.05 / 2.54, 1)
    unmount()
  })
})

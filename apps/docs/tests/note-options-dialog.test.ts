/**
 * BUG-1005: the Footnote/Endnote options dialog must reload the selected
 * kind's numbering when the radio switches. Seeding the fields once from the
 * initially selected kind meant OK silently overwrote the other kind's
 * options (e.g. endnote lowerRoman/start 3) with values the user never saw.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { NoteOptionsDialog } from '../src/renderer/components/NoteOptionsDialog'
import { t } from '../src/renderer/i18n/locale'

describe('NoteOptionsDialog kind switching (BUG-1005)', () => {
  let container: HTMLElement
  let root: Root
  let onApply: ReturnType<typeof vi.fn>
  let onConvert: ReturnType<typeof vi.fn>

  beforeEach(() => {
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    onApply = vi.fn()
    onConvert = vi.fn()
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
  })

  const mount = (value: { footnotes?: never; endnotes?: never } | Record<string, unknown>) =>
    act(() =>
      root.render(
        createElement(NoteOptionsDialog, {
          value,
          hasSelection: false,
          onApply,
          onConvert,
          onClose: vi.fn(),
        }),
      ),
    )

  const radios = () =>
    Array.from(container.querySelectorAll<HTMLInputElement>('input[name="refs-kind"]'))

  const clickOk = () =>
    act(() => {
      const ok = Array.from(container.querySelectorAll('button')).find(
        (b) => b.textContent === t('appOk'),
      )!
      ok.click()
    })

  it('switching to Endnotes re-seeds fmt/start from the endnote options', () => {
    mount({
      footnotes: { numFmt: 'decimal' },
      endnotes: { numFmt: 'lowerRoman', numStart: 3 },
    })
    expect(radios()[0]!.checked).toBe(true) // opens on Footnotes
    act(() => radios()[1]!.click())
    clickOk()
    expect(onApply).toHaveBeenCalledTimes(1)
    // the endnote model comes from the ENDNOTE seed, not the footnote one
    expect(onApply).toHaveBeenCalledWith({
      footnotes: { numFmt: 'decimal' },
      endnotes: { numFmt: 'lowerRoman', numStart: 3 },
    })
  })

  it('a kind without stored options seeds its Word defaults (lowerRoman / 1)', () => {
    mount({ footnotes: { numFmt: 'decimal' } })
    act(() => radios()[1]!.click())
    clickOk()
    expect(onApply).toHaveBeenCalledWith({
      footnotes: { numFmt: 'decimal' },
      endnotes: { numFmt: 'lowerRoman' }, // start 1 is not written
    })
  })

  it('switching back and forth keeps each kind on its own values', () => {
    mount({
      footnotes: { numFmt: 'upperLetter', numStart: 2 },
      endnotes: { numFmt: 'decimal' },
    })
    act(() => radios()[1]!.click()) // → endnotes (decimal/1)
    act(() => radios()[0]!.click()) // → back to footnotes (upperLetter/2)
    clickOk()
    expect(onApply).toHaveBeenCalledWith({
      footnotes: { numFmt: 'upperLetter', numStart: 2 },
      endnotes: { numFmt: 'decimal' },
    })
  })
})

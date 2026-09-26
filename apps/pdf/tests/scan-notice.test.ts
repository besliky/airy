import { beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { ScanNotice } from '../src/renderer/ScanNotice'
import { indexHasNoTextLayer, type SearchIndex } from '../src/renderer/search'

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

/**
 * UX-1733b: a raster-only document gives the user no feedback when they try to
 * select or search text — a quiet status-bar chip explains the situation. The
 * visibility rule is the UX-1733 classifier: the whole index must lack a text
 * layer, and it retires as soon as OCR overlays recognized text anywhere.
 */

const SCAN_INDEX: SearchIndex = [
  { text: '', lower: '', items: [] },
  { text: ' \n ', lower: ' \n ', items: [] },
]

const TEXT_INDEX: SearchIndex = [
  { text: 'Quarterly report', lower: 'quarterly report', items: [] },
  { text: 'Appendix with tables', lower: 'appendix with tables', items: [] },
]

/** What getSearchIndex hands out after the auto-OCR pass recognizes one page */
const OCR_OVERLAID_INDEX: SearchIndex = SCAN_INDEX.map((entry, i) =>
  i === 0 ? { text: 'Recognized page text', lower: 'recognized page text', items: [] } : entry,
)

/** t stub passing the key through, so assertions can pin which key was chosen */
const t = (key: string): string => key

function render(show: boolean, ocr?: boolean): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() =>
    root.render(
      createElement(ScanNotice, {
        show,
        t,
        ...(ocr === undefined ? {} : { ocr }),
      }),
    ),
  )
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

describe('ScanNotice (UX-1733b)', () => {
  it('a document without a text layer shows the affordance', () => {
    const show = indexHasNoTextLayer(SCAN_INDEX)
    expect(show).toBe(true)
    const { container, unmount } = render(show)
    const chip = container.querySelector('.pdf-status-scan')
    expect(chip).not.toBeNull()
    expect(container.querySelector('.pdf-status-scan-label')?.textContent).toBe('scanNoTextLabel')
    // jsdom platform sniff resolves to no OCR — the honest Linux wording
    expect(chip?.getAttribute('data-tip')).toBe('scanNoTextTipNoOcr')
    unmount()
  })

  it('a document with a text layer shows nothing', () => {
    const show = indexHasNoTextLayer(TEXT_INDEX)
    expect(show).toBe(false)
    const { container, unmount } = render(show)
    expect(container.querySelector('.pdf-status-scan')).toBeNull()
    unmount()
  })

  it('retires once OCR has overlaid recognized text (auto-recognition path)', () => {
    expect(indexHasNoTextLayer(OCR_OVERLAID_INDEX)).toBe(false)
    const { container, unmount } = render(false)
    expect(container.querySelector('.pdf-status-scan')).toBeNull()
    unmount()
  })

  it('points to the running OCR path on platforms with an engine', () => {
    const { container, unmount } = render(true, true)
    expect(container.querySelector('.pdf-status-scan')?.getAttribute('data-tip')).toBe(
      'scanNoTextTipOcr',
    )
    unmount()
  })
})

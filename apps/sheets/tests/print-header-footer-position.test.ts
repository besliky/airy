/**
 * BUG-1505 header/footer position: Excel anchors the header/footer text at
 * the file's header/footer margin from the paper edge (LibreOffice measures
 * the top of header text at exactly the header margin), while Chromium
 * prints print-templates inside the margin boxes with a fixed inset of its
 * own. The template's padding compensates so the rendered text lands at
 * Excel's position and stops crossing the content area.
 */
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ BrowserWindow: class {}, dialog: {} }))
vi.mock('@airy-office/electron-utils', () => ({ showSaveDialogWithMemory: vi.fn() }))

import {
  buildHeaderFooterTemplate,
  HEADER_FOOTER_TEMPLATE_INSET_IN,
} from '../src/renderer/print-html'

const PARTS = { left: 'Report' }
const MARGINS = {
  left: 0.7,
  right: 0.7,
  top: 0.5,
  bottom: 0.5,
  header: 0.2,
  footer: 0.3,
}

/// Chromium's template inset as measured on Electron's printToPDF with
/// pdftotext word bboxes: yMin = padding x 72 + inset for every margin,
/// padding and font size tried (14.38..15.13pt across the runs).
const MEASURED_INSET_MIN_PT = 14.38
const MEASURED_INSET_MAX_PT = 15.13
const PT_PER_IN = 72

function anchorPaddingIn(template: string, kind: 'header' | 'footer'): number {
  const anchor = kind === 'header' ? 'padding-top' : 'padding-bottom'
  const match = new RegExp(`${anchor}:([\\d.]+)in`).exec(template)
  expect(match, `${anchor} present in ${template}`).not.toBeNull()
  return Number(match![1])
}

/// Where the rendered text lands: the engine places the template's content
/// at its own fixed inset below the paper edge, padding pushes it further.
function renderedTextTopPt(paddingIn: number): number {
  return paddingIn * PT_PER_IN + HEADER_FOOTER_TEMPLATE_INSET_IN * PT_PER_IN
}

describe('buildHeaderFooterTemplate anchors text at the Excel margin (BUG-1505)', () => {
  const cases: { header: number; footer: number }[] = [
    { header: 0.2, footer: 0.2 },
    { header: 0.25, footer: 0.25 },
    { header: 0.3, footer: 0.3 },
    { header: 0.5, footer: 0.5 },
    { header: 0.75, footer: 0.75 },
  ]

  it.each(cases)(
    'header/footer margin $header/$footer in lands within 1pt of Excel',
    ({ header, footer }) => {
      const headerTemplate = buildHeaderFooterTemplate(
        PARTS,
        'header',
        { ...MARGINS, header, footer },
        'book',
        'Sheet1',
        new Date(2026, 8, 21),
      )
      const footerTemplate = buildHeaderFooterTemplate(
        PARTS,
        'footer',
        { ...MARGINS, header, footer },
        'book',
        'Sheet1',
        new Date(2026, 8, 21),
      )
      expect(headerTemplate).toBeDefined()
      expect(footerTemplate).toBeDefined()
      // Excel: the header text top sits at the header margin from the paper
      // edge; the footer text bottom at the footer margin from the bottom.
      // The compensated inset must cancel the engine's measured inset.
      const headerPadding = anchorPaddingIn(headerTemplate!, 'header')
      expect(Math.abs(renderedTextTopPt(headerPadding) - header * PT_PER_IN)).toBeLessThanOrEqual(1)
      const footerPadding = anchorPaddingIn(footerTemplate!, 'footer')
      expect(Math.abs(renderedTextTopPt(footerPadding) - footer * PT_PER_IN)).toBeLessThanOrEqual(1)
      // The compensation must be exact within the measured inset spread: the
      // padding swallows everything between the smallest and largest measured
      // engine inset, so the residual never exceeds the measurement band.
      expect(header * PT_PER_IN - headerPadding * PT_PER_IN).toBeGreaterThanOrEqual(
        MEASURED_INSET_MIN_PT - 1,
      )
      expect(header * PT_PER_IN - headerPadding * PT_PER_IN).toBeLessThanOrEqual(
        MEASURED_INSET_MAX_PT + 1,
      )
    },
  )

  it('clamps the padding at zero for margins smaller than the engine inset', () => {
    const template = buildHeaderFooterTemplate(
      PARTS,
      'header',
      { ...MARGINS, header: 0.1 },
      'book',
      'Sheet1',
      new Date(2026, 8, 21),
    )
    expect(anchorPaddingIn(template!, 'header')).toBe(0)
  })

  it('keeps the header text clear of the content area (header 0.2 / top 0.5)', () => {
    // Before the fix the padding equalled the header margin, so the text
    // top landed at 14.4pt (margin) + ~15.2pt (engine inset) = 29.6pt and
    // the text bottom crossed the content top at 36pt.
    const template = buildHeaderFooterTemplate(
      PARTS,
      'header',
      MARGINS,
      'book',
      'Sheet1',
      new Date(2026, 8, 21),
    )
    const padding = anchorPaddingIn(template!, 'header')
    const textTopPt = renderedTextTopPt(padding)
    // 9pt text renders in a ~1.25em line box.
    const textBottomPt = textTopPt + 9 * 1.25
    expect(textTopPt).toBeLessThanOrEqual(MARGINS.header * PT_PER_IN + 1)
    expect(textBottomPt).toBeLessThanOrEqual(MARGINS.top * PT_PER_IN)
  })

  it('keeps the side padding on the file margins', () => {
    const template = buildHeaderFooterTemplate(
      PARTS,
      'header',
      MARGINS,
      'book',
      'Sheet1',
      new Date(2026, 8, 21),
    )
    expect(template).toContain(`padding-left:${MARGINS.left}in`)
    expect(template).toContain(`padding-right:${MARGINS.right}in`)
  })
})

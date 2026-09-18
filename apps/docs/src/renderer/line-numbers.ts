// Line numbers (sectPr w:lnNumType): Word's margin numbering of body lines.
// Pure numbering semantics over measured lines + page windows (testable
// without DOM), a DOM line collector, and the canvas overlay painter; the
// pagination preview renders its own copy from the same marks (print/PDF).
import type { LineNumberSettings } from '@airy-office/docx-engine'
import type { BlockBox, PageSlice } from './pagination-types'
import type { DomLineRectsFn } from './pagination-lines'
import { createLineRectsCache } from './pagination-lines'

/** Word's auto distance between the number and the text (twips → px at 96dpi) */
export const LN_AUTO_DISTANCE_PX = 13

/** one counted text line: flow Y in the pagination virtual (gapless) space */
export interface LnLine {
  /** virtual flow top of the line */
  top: number
  /** owning section index (assignSections output; 0 for single-section docs) */
  section: number
  /** line ink height (px, vertical centering of the numeral) */
  height?: number
  /** screen (client) top of the line when sampled on the live canvas */
  screenTop?: number
}

/** one rendered numeral */
export interface LnMark {
  /** slice index (page) the number renders on */
  page: number
  /** regioned pages: region/column the number's column renders in */
  region?: number
  col?: number
  /** Y below the page content top (single-column) / the column top (regions) */
  y: number
  /** the numeral text */
  label: string
  /** owning section index (side/distance come from its settings) */
  section: number
  /** screen (client) top of the line on the live canvas (canvas rendering) */
  screenTop?: number
}

const windowsOfSlice = (
  slice: PageSlice,
): Array<{ start: number; end: number; region: number; col: number }> => {
  if (!slice.regions) return [{ start: slice.start, end: slice.end, region: -1, col: -1 }]
  const out: Array<{ start: number; end: number; region: number; col: number }> = []
  slice.regions.forEach((r, ri) =>
    r.columns.forEach((c, ci) => out.push({ start: c.start, end: c.end, region: ri, col: ci })),
  )
  return out
}

/**
 * Word's line numbering over one pagination result. Lines are counted in
 * virtual flow order (column order on multi-column pages): Word numbers every
 * body line, restarting the counter at `start` on each restart point; a
 * numeral displays on lines whose count is a multiple of countBy (countBy 5
 * shows 5, 10, 15… — the first line's value can stay unshown).
 *
 * Sections without w:lnNumType neither display nor count lines; `continuous`
 * keeps one running counter across pages and numbered sections, while
 * newPage/newSection reset it to that section's start.
 */
export function computeLineNumberMarks(
  lines: LnLine[],
  slices: PageSlice[],
  sections: Array<{ settings?: { lineNumbers?: LineNumberSettings } }>,
): LnMark[] {
  if (!sections.some((s) => s.settings?.lineNumbers)) return []
  const sorted = [...lines].sort((a, b) => a.top - b.top)
  // per-slice column windows in reading order (page → region → column)
  const windows = slices.map((s) => ({ section: s.section, wins: windowsOfSlice(s) }))
  const findWindow = (top: number) => {
    for (let p = 0; p < windows.length; p++) {
      for (const w of windows[p].wins) {
        if (top >= w.start - 0.5 && top < w.end - 0.5) return { page: p, win: w }
      }
    }
    return null
  }
  const marks: LnMark[] = []
  /** running counter state; sectionIdx = -1 before any numbered section */
  let counter = 0
  let counterSection = -1
  let lastNumberedPage = -1
  for (const line of sorted) {
    const sec = sections[Math.min(line.section, sections.length - 1)]
    const ln = sec?.settings?.lineNumbers
    if (!ln) continue
    const start = ln.start ?? 1
    const countBy = ln.countBy ?? 1
    const restart = ln.restart ?? 'newPage'
    const loc = findWindow(line.top)
    if (!loc) continue
    const newSection = line.section !== counterSection
    const newPage = loc.page !== lastNumberedPage
    if (counterSection < 0) counter = start
    else if (restart === 'newPage' && (newPage || newSection)) counter = start
    else if (restart === 'newSection' && newSection) counter = start
    // continuous: the running counter survives pages and section breaks
    counterSection = line.section
    lastNumberedPage = loc.page
    const label = counter
    counter++
    if (label % countBy === 0) {
      marks.push({
        page: loc.page,
        ...(loc.win.region >= 0 ? { region: loc.win.region, col: loc.win.col } : {}),
        y: line.top - loc.win.start,
        label: String(label),
        section: line.section,
        ...(line.screenTop !== undefined ? { screenTop: line.screenTop } : {}),
      })
    }
  }
  return marks
}

/**
 * Word does not number lines inside tables, text boxes or floating objects;
 * empty paragraphs still count as one line each.
 */
function blockCountsLines(b: BlockBox): boolean {
  if (b.floated || b.isEndnotes || b.isFloatSpill) return false
  const el = b.el
  if (!el) return false
  if (
    el.classList.contains('doc-protected-textboxes') ||
    el.classList.contains('doc-protected-image') ||
    el.classList.contains('img-wrap-band')
  )
    return false
  // tables (native or nested in a protected passthrough) are skipped like Word
  if (b.tableRows || el.querySelector('tr')) return false
  return true
}

/** measured lines of one block: DOM line rects, or the block itself when empty */
export function collectBlockLines(
  block: BlockBox,
  rectsOf: DomLineRectsFn,
  zoomFactor: number,
): LnLine[] {
  if (!blockCountsLines(block) || !block.el) return []
  const lines: LnLine[] = []
  for (const ln of rectsOf(block.el, zoomFactor)) {
    lines.push({
      top: block.top + ln.offset,
      section: block.section ?? 0,
      ...(ln.bottom - ln.offset > 0 ? { height: ln.bottom - ln.offset } : {}),
      screenTop: ln.top,
    })
  }
  if (lines.length === 0 && block.height > 1) {
    // empty paragraph mark: Word counts the blank line
    const r = block.el.getBoundingClientRect()
    lines.push({
      top: block.top,
      section: block.section ?? 0,
      height: block.height,
      screenTop: r.top,
    })
  }
  return lines
}

/** section-settings shape the canvas painter reads (SectionSettings satisfies it) */
export type LnSectionSet = {
  pageWidth: number
  marginLeft: number
  marginRight: number
  bidi?: boolean
  lineNumbers?: LineNumberSettings
}

/**
 * Canvas overlay: one absolutely positioned numeral per mark at the line's
 * screen position, right-aligned to the text edge minus the distance (left
 * margin of LTR sections, right margin of w:bidi sections). Unzoomed wrap
 * coordinates, like syncPageBorders; must run after the page gaps are placed.
 */
export function syncLineNumberOverlays(
  wrap: HTMLElement,
  blocks: BlockBox[],
  slices: PageSlice[],
  sections: Array<{ settings?: LnSectionSet }>,
  zoomFactor: number,
  rectsOf: DomLineRectsFn = createLineRectsCache(),
): void {
  const anyNumbering = sections.some((s) => s.settings?.lineNumbers)
  let layer = wrap.querySelector(':scope > .page-linenum-overlays') as HTMLElement | null
  if (!anyNumbering) {
    layer?.remove()
    return
  }
  const lines: LnLine[] = []
  for (const b of blocks) for (const ln of collectBlockLines(b, rectsOf, zoomFactor)) lines.push(ln)
  const marks = computeLineNumberMarks(lines, slices, sections)
  if (marks.length === 0) {
    layer?.remove()
    return
  }
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-linenum-overlays'
    wrap.appendChild(layer)
  }
  layer.textContent = ''
  const wrapTop = wrap.getBoundingClientRect().top
  for (const m of marks) {
    const set = sections[Math.min(m.section, sections.length - 1)].settings
    const ln = set?.lineNumbers
    if (!ln || m.screenTop === undefined) continue
    const dist = ln.distance !== undefined ? lnTwipsToPx(ln.distance) : LN_AUTO_DISTANCE_PX
    const el = document.createElement('div')
    el.className = 'page-linenum'
    el.textContent = m.label
    el.style.top = `${((m.screenTop - wrapTop) / zoomFactor).toFixed(1)}px`
    if (set.bidi === true) {
      const left = lnTwipsToPx(set.pageWidth - set.marginRight) + dist
      el.style.left = `${left.toFixed(1)}px`
    } else {
      const right = lnTwipsToPx(set.pageWidth - set.marginLeft) + dist
      el.style.right = `${right.toFixed(1)}px`
    }
    layer.appendChild(el)
  }
}

/** twips → CSS px at 96dpi (self-contained copy of the renderer's twipsToPx) */
export const lnTwipsToPx = (twips: number) => (twips / 1440) * 96

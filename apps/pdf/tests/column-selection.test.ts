import { describe, expect, it } from 'vitest'
import {
  clusterLineBoxes,
  columnOfBox,
  computeSelectionBand,
  detectColumns,
  distBoxToSegment,
  planBandCorrection,
  type ColumnLayout,
  type LineItem,
  type SpanBox,
} from '../src/renderer/column-selection'

// Synthetic page geometry (client px): A4-ish width, two 240px columns with a 40px
// gutter, 26 rows of five 40px "words" per column. Mirrors the audited twocol.pdf
// repro (PDF-UX-1724 / BUG-1729): drag from left column row 0 to right column row 2.
const ROW_H = 12
const ROW_STEP = 14
const ROWS = 26
const COLS = [40, 320]
const COL_W = 240

const box = (left: number, top: number, width: number, height: number): SpanBox => ({
  left,
  top,
  right: left + width,
  bottom: top + height,
})

/** A word box in column `col` (0 = left), row `r`, word `w`. */
const wordBox = (col: number, r: number, w: number) =>
  box(COLS[col]! + w * 48, 20 + r * ROW_STEP, 40, ROW_H)

interface Word {
  id: string
  box: SpanBox
}

const rowY = (r: number) => 20 + r * ROW_STEP + ROW_H / 2

function twoColumnWords(rows = ROWS): Word[] {
  const words: Word[] = []
  for (let r = 0; r < rows; r++)
    for (let col = 0; col < 2; col++)
      for (let w = 0; w < 5; w++) words.push({ id: `c${col}r${r}w${w}`, box: wordBox(col, r, w) })
  return words
}

function singleColumnWords(rows = ROWS): Word[] {
  const words: Word[] = []
  for (let r = 0; r < rows; r++)
    for (let w = 0; w < 11; w++)
      words.push({ id: `r${r}w${w}`, box: box(40 + w * 48, 20 + r * ROW_STEP, 40, ROW_H) })
  return words
}

const boxOf = (word: Word) => word.box

const asLines = (words: Word[]): LineItem<Word>[] => clusterLineBoxes(words, boxOf)

/** Selection client rects covering whole boxes (what a native selection paints). */
const rectsOf = (words: Word[]): SpanBox[] => words.map((w) => w.box)

const ALL_LAYOUT: ColumnLayout = { gutters: [{ left: 280, right: 320 }] }

describe('clusterLineBoxes', () => {
  it('merges same-row words across both columns into one visual line', () => {
    const lines = asLines(twoColumnWords())
    expect(lines).toHaveLength(ROWS)
    // Words inside a line are ordered left to right
    expect(lines[0]!.items[0]!.id).toBe('c0r0w0')
    expect(lines[0]!.items[5]!.id).toBe('c1r0w0')
    expect(lines[0]!.box).toEqual({ left: 40, top: 20, right: 552, bottom: 32 })
  })

  it('keeps lines with staggered baselines apart when overlap is under half the height', () => {
    const a = { id: 'a', box: box(0, 0, 100, 10) }
    const b = { id: 'b', box: box(0, 8, 100, 10) }
    expect(asLines([a, b])).toHaveLength(2)
  })
})

describe('detectColumns', () => {
  it('finds the single gutter of a two-column page', () => {
    const layout = detectColumns(twoColumnWords(), boxOf)
    expect(layout).not.toBeNull()
    expect(layout!.gutters).toHaveLength(1)
    const g = layout!.gutters[0]!
    // The gutter spans the free space between the columns' word boxes (272..320)
    expect(g.left).toBeGreaterThanOrEqual(272)
    expect(g.left).toBeLessThan(320)
    expect(g.right).toBeLessThanOrEqual(320)
    expect(g.right).toBeGreaterThan(272)
    // Columns read left then right
    expect(columnOfBox(wordBox(0, 0, 0), layout!)).toBe(0)
    expect(columnOfBox(wordBox(1, 0, 0), layout!)).toBe(1)
  })

  it('keeps the single-column page undetected (previous behavior)', () => {
    expect(detectColumns(singleColumnWords(), boxOf)).toBeNull()
  })

  it('refuses a layout when a full-width title crosses the gutter', () => {
    const title = { id: 'title', box: box(40, 0, 520, ROW_H) }
    const body = twoColumnWords().map((w) => ({
      ...w,
      box: { ...w.box, top: w.box.top + 28, bottom: w.box.bottom + 28 },
    }))
    expect(detectColumns([title, ...body], boxOf)).toBeNull()
  })

  it('finds two gutters on a three-column page', () => {
    const words: Word[] = []
    for (let r = 0; r < 10; r++)
      for (let col = 0; col < 3; col++)
        for (let w = 0; w < 3; w++)
          words.push({
            id: `c${col}r${r}w${w}`,
            box: box(40 + col * 190 + w * 48, 20 + r * ROW_STEP, 40, ROW_H),
          })
    const layout = detectColumns(words, boxOf)
    expect(layout).not.toBeNull()
    expect(layout!.gutters).toHaveLength(2)
  })

  it('needs enough lines before trusting a layout', () => {
    expect(detectColumns(twoColumnWords(3), boxOf)).toBeNull()
  })
})

describe('distBoxToSegment', () => {
  it('is zero when the segment crosses the box', () => {
    expect(distBoxToSegment(box(0, 0, 10, 10), { x: -5, y: 5 }, { x: 15, y: 5 })).toBe(0)
  })

  it('measures the gap for a segment passing near the box', () => {
    expect(distBoxToSegment(box(0, 20, 100, 30), { x: 0, y: 0 }, { x: 100, y: 0 })).toBe(20)
  })
})

describe('computeSelectionBand', () => {
  const lines = asLines(twoColumnWords())
  // The audited repro: drag from "ALPHA left 0" to "BETA right 2"
  const anchor = { x: COLS[0]!, y: rowY(0) }
  const focus = { x: COLS[1]! + 20, y: rowY(2) }

  it('selects the swept rows of both columns, not a whole column', () => {
    const band = computeSelectionBand(lines, ALL_LAYOUT, anchor, focus, boxOf)
    expect(band).not.toBeNull()
    const bandRows = new Set(band!.map((c) => c.line)).size
    expect(bandRows).toBeLessThanOrEqual(5)
    // Nothing from the lower half of the page may leak in
    const maxRow = Math.max(...band!.map((c) => Math.round((c.line.box.top - 20) / ROW_STEP)))
    expect(maxRow).toBeLessThanOrEqual(4)
    expect(maxRow).toBeGreaterThanOrEqual(0)
  })

  it('orders the band for reading: left column rows first, then right column rows', () => {
    const band = computeSelectionBand(lines, ALL_LAYOUT, anchor, focus, boxOf)!
    const cols = band.map((c) => columnOfBox(c.box, ALL_LAYOUT))
    // Non-decreasing column index: all column-0 cells precede column-1 cells
    expect([...cols].sort((a, b) => a - b)).toEqual(cols)
    const firstRight = cols.indexOf(1)
    expect(cols.slice(0, firstRight).every((c) => c === 0)).toBe(true)
  })

  it('returns null when both endpoints share a column (native selection is fine)', () => {
    expect(
      computeSelectionBand(lines, ALL_LAYOUT, { x: 50, y: rowY(0) }, { x: 60, y: rowY(5) }, boxOf),
    ).toBeNull()
  })
})

describe('planBandCorrection', () => {
  const words = twoColumnWords()
  const lines = asLines(words)
  const base = {
    layout: ALL_LAYOUT,
    lines,
    boxOf,
  }

  it('rewrites the audited over-selection (whole left column) into the band', () => {
    // Native DOM-order selection painted every row (the whole left column path)
    const plan = planBandCorrection({
      ...base,
      selectionRects: rectsOf(words),
      anchor: { x: COLS[0]!, y: rowY(0) },
      focus: { x: COLS[1]! + 20, y: rowY(2) },
    })
    expect(plan).not.toBeNull()
    const bandRows = new Set(plan!.map((c) => c.line)).size
    expect(bandRows).toBeLessThanOrEqual(5)
    expect(plan!.length).toBeGreaterThan(bandRows) // both columns contribute cells
  })

  it('leaves a selection with no over-reach to the native behavior', () => {
    // Only the band rows themselves are selected: nothing to correct
    const bandWords = words.filter((w) => {
      const r = Math.round((w.box.top - 20) / ROW_STEP)
      return r <= 2
    })
    const plan = planBandCorrection({
      ...base,
      selectionRects: rectsOf(bandWords),
      anchor: { x: COLS[0]!, y: rowY(0) },
      focus: { x: COLS[1]! + 20, y: rowY(2) },
    })
    expect(plan).toBeNull()
  })

  it('never rewrites a plain select-all', () => {
    const plan = planBandCorrection({
      ...base,
      selectionRects: rectsOf(words),
      anchor: { x: COLS[0]!, y: rowY(0) },
      focus: { x: COLS[1]! + COL_W, y: rowY(ROWS - 1) },
    })
    expect(plan).toBeNull()
  })

  it('does nothing without a detected column layout (single-column pages)', () => {
    const single = singleColumnWords()
    const plan = planBandCorrection({
      layout: null,
      lines: asLines(single),
      selectionRects: rectsOf(single),
      anchor: { x: 50, y: rowY(0) },
      focus: { x: 500, y: rowY(2) },
      boxOf,
    })
    expect(plan).toBeNull()
  })
})

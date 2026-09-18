import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import type { LineNumberSettings } from '@airy-office/docx-engine'
import {
  collectBlockLines,
  computeLineNumberMarks,
  LN_AUTO_DISTANCE_PX,
  lnTwipsToPx,
  syncLineNumberOverlays,
  type LnLine,
} from '../src/renderer/line-numbers'
import type { BlockBox, PageSlice } from '../src/renderer/pagination-types'
import type { DomLineRectsFn } from '../src/renderer/pagination-lines'

const line = (top: number, section = 0): LnLine => ({ top, section })
const slice = (start: number, end: number, section = 0): PageSlice => ({ start, end, section })

const secs = (ln: LineNumberSettings) => [{ settings: { lineNumbers: ln } }]

describe('computeLineNumberMarks semantics', () => {
  it('numbers every body line by default', () => {
    const marks = computeLineNumberMarks([line(0), line(20), line(40)], [slice(0, 1000)], secs({}))
    expect(marks.map((m) => [m.page, m.label, m.y])).toEqual([
      [0, '1', 0],
      [0, '2', 20],
      [0, '3', 40],
    ])
  })

  it('countBy shows only every Nth line, labelled with its line number', () => {
    const lines = [0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220].map((t) => line(t))
    const marks = computeLineNumberMarks(lines, [slice(0, 1000)], secs({ countBy: 5 }))
    expect(marks.map((m) => m.label)).toEqual(['5', '10'])
    expect(marks.map((m) => m.y)).toEqual([80, 180])
  })

  it('start shifts the counter (start 3, count by 2 → 4, 6)', () => {
    const lines = [0, 20, 40, 60].map((t) => line(t))
    const marks = computeLineNumberMarks(lines, [slice(0, 1000)], secs({ start: 3, countBy: 2 }))
    expect(marks.map((m) => m.label)).toEqual(['4', '6'])
  })

  it('restart newPage (the default) restarts numbering on each page', () => {
    const lines = [0, 20, 500, 520].map((t) => line(t))
    const marks = computeLineNumberMarks(
      lines,
      [slice(0, 300), slice(300, 1000)],
      secs({ restart: 'newPage' }),
    )
    expect(marks.map((m) => [m.page, m.label])).toEqual([
      [0, '1'],
      [0, '2'],
      [1, '1'],
      [1, '2'],
    ])
  })

  it('restart newSection restarts at the section boundary, not each page', () => {
    const lines = [line(0, 0), line(20, 0), line(1500, 1), line(2000, 1)]
    const marks = computeLineNumberMarks(
      lines,
      [slice(0, 1000, 0), slice(1000, 3000, 1)],
      secs({ restart: 'newSection' }),
    )
    expect(marks.map((m) => [m.page, m.label])).toEqual([
      [0, '1'],
      [0, '2'],
      [1, '1'],
      [1, '2'],
    ])
  })

  it('continuous keeps one counter across pages and sections', () => {
    const lines = [line(0, 0), line(500, 0), line(1500, 1), line(2500, 1)]
    const marks = computeLineNumberMarks(
      lines,
      [slice(0, 1000, 0), slice(1000, 2000, 1), slice(2000, 3000, 1)],
      secs({ restart: 'continuous' }),
    )
    expect(marks.map((m) => [m.page, m.label])).toEqual([
      [0, '1'],
      [0, '2'],
      [1, '3'],
      [2, '4'],
    ])
  })

  it('sections without w:lnNumType neither display nor count lines', () => {
    const lines = [line(0, 0), line(500, 0), line(1500, 1), line(2500, 2)]
    const sections: Array<{ settings?: { lineNumbers?: LineNumberSettings } }> = [
      { settings: { lineNumbers: { restart: 'continuous' } } },
      { settings: {} },
      { settings: { lineNumbers: { restart: 'continuous' } } },
    ]
    const marks = computeLineNumberMarks(
      lines,
      [slice(0, 1000, 0), slice(1000, 2000, 1), slice(2000, 3000, 2)],
      sections,
    )
    expect(marks.map((m) => [m.page, m.label])).toEqual([
      [0, '1'],
      [0, '2'],
      // continuous keeps counting past the unnumbered middle section
      [2, '3'],
    ])
  })

  it('multi-column pages number each column, continuing the count', () => {
    const lines = [line(0), line(20), line(500), line(520)]
    const marks = computeLineNumberMarks(
      lines,
      [
        {
          start: 0,
          end: 1040,
          section: 0,
          regions: [
            {
              top: 0,
              height: 500,
              section: 0,
              columns: [
                { start: 0, end: 500 },
                { start: 500, end: 1000 },
              ],
            },
          ],
        },
      ],
      secs({}),
    )
    expect(marks.map((m) => [m.col, m.label, m.y])).toEqual([
      [0, '1', 0],
      [0, '2', 20],
      [1, '3', 0],
      [1, '4', 20],
    ])
  })

  it('no numbered section = no marks', () => {
    expect(computeLineNumberMarks([line(0)], [slice(0, 100)], [{ settings: {} }])).toEqual([])
  })
})

const fakeRects =
  (tops: number[], el: HTMLElement): DomLineRectsFn =>
  (target) => {
    if (target !== el) return []
    return tops.map((top, i) => ({
      offset: top,
      bottom: top + 12,
      left: 0,
      top: 700 + i * 20,
      node: document.createTextNode('x'),
    }))
  }

describe('collectBlockLines (Word counting rules)', () => {
  const blockWith = (over: Partial<BlockBox> = {}): BlockBox => {
    const el = document.createElement('p')
    el.textContent = 'text'
    return { top: 100, height: 60, el, section: 0, ...over }
  }

  it('maps DOM line rects into virtual lines with screen tops', () => {
    const b = blockWith({})
    const lines = collectBlockLines(b, fakeRects([0, 24], b.el!), 1)
    expect(lines.map((l) => l.top)).toEqual([100, 124])
    expect(lines.map((l) => l.screenTop)).toEqual([700, 720])
  })

  it('empty paragraphs count as one line (Word counts blank lines)', () => {
    const b = blockWith({ height: 24 })
    const lines = collectBlockLines(b, fakeRects([], b.el!), 1)
    expect(lines).toEqual([{ top: 100, section: 0, height: 24, screenTop: 0 }])
  })

  it('tables, textboxes, images, endnotes and floats are skipped', () => {
    const table = document.createElement('table')
    table.innerHTML = '<tr><td>x</td></tr>'
    const cases: Array<Partial<BlockBox>> = [
      { tableRows: [{ height: 10 }] },
      { el: table },
      { floated: true },
      { isEndnotes: true },
      { isFloatSpill: true },
    ]
    for (const over of cases) {
      const b = { ...blockWith(over), ...over }
      expect(collectBlockLines(b, fakeRects([0], b.el!), 1)).toEqual([])
    }
    const tb = blockWith({})
    tb.el!.className = 'doc-protected-textboxes'
    expect(collectBlockLines(tb, fakeRects([0], tb.el!), 1)).toEqual([])
  })
})

describe('syncLineNumberOverlays (canvas)', () => {
  const settings = {
    pageWidth: 11906,
    pageHeight: 16838,
    marginLeft: 1440,
    marginRight: 1440,
    lineNumbers: { restart: 'newPage' as const },
  }

  const block = (): BlockBox => {
    const el = document.createElement('p')
    el.textContent = 'hello world'
    return { top: 0, height: 40, section: 0, el }
  }

  it('paints one positioned numeral per mark and clears when numbering is off', () => {
    const wrap = document.createElement('div')
    const b = block()
    // two fake DOM lines → two numerals
    syncLineNumberOverlays(
      wrap,
      [b],
      [slice(0, 1000)],
      [{ settings }],
      1,
      fakeRects([0, 20], b.el!),
    )
    const nums = wrap.querySelectorAll<HTMLElement>('.page-linenum')
    expect([...nums].map((n) => n.textContent)).toEqual(['1', '2'])
    // numbering off → the whole layer is removed
    const wrap2 = document.createElement('div')
    const layer = document.createElement('div')
    layer.className = 'page-linenum-overlays'
    wrap2.appendChild(layer)
    syncLineNumberOverlays(
      wrap2,
      [block()],
      [slice(0, 1000)],
      [{ settings: { ...settings, lineNumbers: undefined } }],
      1,
    )
    expect(wrap2.querySelector('.page-linenum-overlays')).toBeNull()
  })

  it('distance and margins place the numeral right of the text edge (twips → px)', () => {
    const b = block()
    const wrap = document.createElement('div')
    const set = { ...settings, lineNumbers: { restart: 'newPage' as const, distance: 240 } }
    syncLineNumberOverlays(
      wrap,
      [b],
      [slice(0, 1000)],
      [{ settings: set }],
      1,
      fakeRects([0], b.el!),
    )
    const num = wrap.querySelector<HTMLElement>('.page-linenum')
    expect(num?.style.right).toBe(`${(lnTwipsToPx(11906 - 1440) + lnTwipsToPx(240)).toFixed(1)}px`)
    // screen top 700 (fake rect) minus wrap top 0, factor 1 (jsdom drops the .0)
    expect(num?.style.top).toBe('700px')
    // auto distance when w:distance is absent
    const wrapAuto = document.createElement('div')
    const b2 = block()
    syncLineNumberOverlays(
      wrapAuto,
      [b2],
      [slice(0, 1000)],
      [{ settings }],
      1,
      fakeRects([0], b2.el!),
    )
    const autoNum = wrapAuto.querySelector<HTMLElement>('.page-linenum')
    expect(autoNum?.style.right).toBe(
      `${(lnTwipsToPx(11906 - 1440) + LN_AUTO_DISTANCE_PX).toFixed(1)}px`,
    )
  })
})

describe('canvas numeral ink and font (UX-901: document data, not chrome)', () => {
  const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')

  /** declarations of the first rule whose selector matches (styles.css contract) */
  const ruleDecls = (selector: string): ReadonlyMap<string, string> => {
    const m = new RegExp(`^${selector.replace(/[.]/g, '\\$&')} \\{([^}]*)\\}`, 'm').exec(css)
    if (!m) throw new Error(`rule not found: ${selector}`)
    const out = new Map<string, string>()
    for (const d of m[1].matchAll(/([\w-]+):\s*([^;]+);/g)) out.set(d[1], d[2].trim())
    return out
  }

  /** body of the first rule matching the selector prefix (brace-matched) */
  const blockOf = (selector: string): string => {
    const i = css.indexOf(selector)
    if (i < 0) throw new Error(`selector not found: ${selector}`)
    const open = css.indexOf('{', i)
    let depth = 0
    for (let j = open; j < css.length; j++) {
      if (css[j] === '{') depth++
      else if (css[j] === '}' && --depth === 0) return css.slice(open + 1, j)
    }
    throw new Error(`unbalanced block: ${selector}`)
  }

  it('paints numerals in paper ink, never the chrome --text token', () => {
    const num = ruleDecls('.page-linenum')
    // the overlay layer hangs off .page-wrap — a sibling of .doc-page — so the
    // page's inherited ink never reaches the numerals; the rule must carry it
    expect(num.get('color')).toBe('var(--docs-paper-ink)')
    expect([...num.values()].join(' ')).not.toContain('var(--text')
  })

  it('uses the document font of the print copy (canvas == print)', () => {
    const num = ruleDecls('.page-linenum')
    const pv = ruleDecls('.pv-page')
    for (const prop of ['font-family', 'font-size', 'line-height']) {
      expect(num.get(prop), prop).toBe(pv.get(prop))
    }
  })

  it('the ink follows the page, not the UI theme (all four theme/page combos)', () => {
    // light paper: black numerals in BOTH UI themes…
    expect(blockOf(':root,')).toContain('--docs-paper-ink: #000;')
    const uiDark = blockOf("[data-theme='dark']")
    expect(uiDark, 'dark UI theme must not restyle document ink').not.toContain('--docs-paper-ink')
    // …dark page (View ▸ Dark Mode): white numerals in both UI themes
    expect(blockOf('.page-dark {')).toContain('--docs-paper-ink: #ffffff;')
  })
})

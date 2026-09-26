import { describe, expect, it } from 'vitest'
import { buildPrintDocumentHtml, parsePrintRange, printPageCount } from '../src/shared/print-html'

const srcs = (n: number) => Array.from({ length: n }, (_x, i) => `blob:img-${i}`)

describe('parsePrintRange', () => {
  it('parses singles, ranges and mixed separators, deduped and sorted', () => {
    expect(parsePrintRange('1,3,5-8', 10)).toEqual([0, 2, 4, 5, 6, 7])
    expect(parsePrintRange('3，1、2 5;4', 5)).toEqual([0, 1, 2, 3, 4])
    expect(parsePrintRange('2-4, 3', 5)).toEqual([1, 2, 3])
  })

  it('rejects empty, malformed, out-of-bounds and inverted input', () => {
    expect(parsePrintRange('', 5)).toBeNull()
    expect(parsePrintRange('a-b', 5)).toBeNull()
    expect(parsePrintRange('0-2', 5)).toBeNull()
    expect(parsePrintRange('4-2', 5)).toBeNull()
    expect(parsePrintRange('6', 5)).toBeNull()
  })
})

describe('printPageCount', () => {
  it('follows the per-page thumbnail count of each layout', () => {
    expect(printPageCount(7, 'full')).toBe(7)
    expect(printPageCount(7, 'notes')).toBe(7)
    expect(printPageCount(7, 'handout2')).toBe(4)
    expect(printPageCount(7, 'handout3')).toBe(3)
    expect(printPageCount(7, 'handout6')).toBe(2)
  })
})

describe('buildPrintDocumentHtml svgs', () => {
  it('inline vector slides replace their <img> slot without changing the layout', () => {
    const svgs = [
      '<svg xmlns="http://www.w3.org/2000/svg"><text>First</text></svg>',
      undefined,
      '<svg xmlns="http://www.w3.org/2000/svg"><text>Third</text></svg>',
    ]
    const html = buildPrintDocumentHtml({
      srcs: srcs(3),
      svgs,
      ratio: 16 / 9,
      layout: 'handout2',
    })
    // 3 slides at 2 per sheet = 2 A4 pages, one bitmap slot kept
    expect(html.match(/class="page handout h2"/g)).toHaveLength(2)
    expect(html).toContain('<svg xmlns="http://www.w3.org/2000/svg"><text>First</text></svg>')
    expect(html).toContain('<div class="cell"><img src="blob:img-1">')
    // the svg cells carry the same frame rules as the bitmap cells
    expect(html).toContain('.page.handout .cell svg { border: 1px solid #bbb;')
  })

  it('notes pages inline the svg above the escaped notes text', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      svgs: ['<svg><text>Deck</text></svg>'],
      ratio: 16 / 9,
      layout: 'notes',
      notes: ['Say <this>'],
    })
    expect(html).toMatch(
      /<div class="page notes"><svg><text>Deck<\/text><\/svg><div class="note">Say &lt;this&gt;<\/div><\/div>/,
    )
  })
})

describe('buildPrintDocumentHtml', () => {
  it('full layout sizes pages by the slide ratio and can frame slides', () => {
    const html = buildPrintDocumentHtml({ srcs: srcs(2), ratio: 16 / 9, layout: 'full' })
    expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    expect(html.match(/class="page"/g)).toHaveLength(2)
    expect(html).not.toContain('.page > img { border')

    const framed = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'full',
      frame: true,
    })
    expect(framed).toContain('.page > img { border: 1px solid #bbb; }')
  })

  it('handouts use A4 and swap dimensions in landscape', () => {
    const portrait = buildPrintDocumentHtml({ srcs: srcs(5), ratio: 16 / 9, layout: 'handout6' })
    expect(portrait).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
    expect(portrait).toContain('grid-template-columns: 1fr 1fr;')

    const landscape = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout6',
      orientation: 'landscape',
    })
    expect(landscape).toContain('@page { size: 11.69in 8.27in; margin: 0; }')
    expect(landscape).toContain('grid-template-columns: 1fr 1fr 1fr;')
  })

  it('notes layout escapes note text and pairs it with each slide', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(2),
      ratio: 4 / 3,
      layout: 'notes',
      notes: ['a < b\nnext', ''],
    })
    expect(html.match(/class="page notes"/g)).toHaveLength(2)
    expect(html).toContain('a &lt; b<br>next')
  })

  it('notes pages are not clipping boxes: they grow and fragment onto continuation sheets', () => {
    const html = buildPrintDocumentHtml({ srcs: srcs(1), ratio: 16 / 9, layout: 'notes' })
    // the generic .page clip (fixed height + overflow: hidden) must be overridden
    expect(html).toContain(
      '.page.notes { padding: 0.5in; height: auto; min-height: 11.69in; overflow: visible;',
    )
    // block flow: flex items do not fragment across printed pages
    expect(html.match(/\.page\.notes \{[^}]*display: flex[^}]*\}/g)).toBeNull()
  })

  it('BUG-1766: a 30k-character note is fully in the printed flow, not silently clipped', () => {
    const note30k = 'LOREMNOTE'.repeat(3000) // the audit corpus shape: 30 000 chars, one line
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notes: [note30k],
    })
    // every character survives into the DOM (the PDF text layer)
    expect(html).toContain(note30k)
    // still one page block per slide: the overflow continues via pagination,
    // not via extra slide-less divs
    expect(html.match(/class="page notes"/g)).toHaveLength(1)

    const lines = Array.from({ length: 400 }, (_x, i) => `Line ${i}: lorem`).join('\n')
    const lineHtml = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notes: [lines],
    })
    expect(lineHtml).toContain('Line 0: lorem<br>Line 1: lorem')
    expect(lineHtml).toContain('Line 399: lorem')
  })

  it('short notes keep the single framed page and the 11pt note styles', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notes: ['tiny'],
    })
    expect(html).toContain(
      '<div class="page notes"><img src="blob:img-0"><div class="note">tiny</div></div>',
    )
    expect(html).toContain(
      '.page.notes .note { margin-top: 0.3in; font-size: 11pt; line-height: 1.5; white-space: pre-wrap; overflow-wrap: anywhere; }',
    )
  })

  it('landscape notes bound the slide image with an explicit box (no percentage heights on auto pages)', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      orientation: 'landscape',
      notes: ['n'],
    })
    expect(html).toContain('min-height: 8.27in')
    // 55% of the landscape content height (7.27in), bounded by the inner width, at the slide ratio
    expect(html).toContain('width: 7.108in; height: 3.999in; margin: 0 auto;')
  })

  it('preview mode adds page badges with the total page count', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout2',
      preview: true,
    })
    expect(html).toContain("content: counter(pg) ' / 3';")
    expect(html).toContain('counter-reset: pg;')
  })
})

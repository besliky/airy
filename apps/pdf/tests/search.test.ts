import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  buildSearchIndex,
  MAX_MATCHES,
  nextMatchIndex,
  searchInIndex,
  type SearchIndex,
} from '../src/renderer/search'

interface FakeItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
}

function fakeDoc(pages: FakeItem[][]): PDFDocumentProxy {
  return {
    numPages: pages.length,
    getPage: async (n: number) => ({
      getTextContent: async () => ({ items: pages[n - 1] }),
    }),
  } as unknown as PDFDocumentProxy
}

const item = (str: string, x: number, y: number, w: number, h: number): FakeItem => ({
  str,
  transform: [1, 0, 0, 1, x, y],
  width: w,
  height: h,
})

describe('buildSearchIndex', () => {
  it('concatenates item text per page and records char ranges', async () => {
    const doc = fakeDoc([[item('Hello ', 10, 700, 60, 12), item('World', 70, 700, 50, 12)]])
    const index = await buildSearchIndex(doc)
    expect(index).toHaveLength(1)
    expect(index[0]!.text).toBe('Hello World')
    expect(index[0]!.lower).toBe('hello world')
    expect(index[0]!.items).toEqual([
      { start: 0, end: 6, x: 10, y: 700, w: 60, h: 12 },
      { start: 6, end: 11, x: 70, y: 700, w: 50, h: 12 },
    ])
  })

  it('inserts newlines for hasEOL and skips empty/invalid items', async () => {
    const doc = fakeDoc([
      [
        { ...item('line1', 0, 0, 10, 10), hasEOL: true },
        { str: '', hasEOL: true }, // empty text still contributes the EOL
        { transform: [1, 0, 0, 1, 0, 0] }, // no str -> skipped entirely
        item('line2', 0, 0, 10, 10),
      ],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.text).toBe('line1\n\nline2')
  })

  it('derives height from the transform when height is missing', async () => {
    const doc = fakeDoc([[{ str: 'x', transform: [1, 0, 3, 4, 0, 0], width: 5 }]])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.h).toBe(5) // hypot(3, 4)
  })

  it('flags rotated runs and leaves upright ones unflagged', async () => {
    const doc = fakeDoc([
      [
        item('upright', 10, 700, 60, 12),
        { str: 'rotated', transform: [0, 12, -12, 0, 200, 400], width: 60, height: 12 },
      ],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.rot).toBeUndefined()
    expect(index[0]!.items[1]!.rot).toBe(true)
  })

  it('synthetic italic shear (c ≠ 0, horizontal baseline) is not flagged as rotated', async () => {
    // ~12° shear as writers emit for fake italics: b = 0, c = tan(12°) × size
    const doc = fakeDoc([
      [{ str: 'emphasis', transform: [12, 0, 2.55, 12, 100, 700], width: 48, height: 12 }],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.rot).toBeUndefined()
  })
})

describe('searchInIndex', () => {
  const entry = (text: string, items: SearchIndex[number]['items']): SearchIndex[number] => ({
    text,
    lower: text.toLowerCase(),
    items,
  })

  it('returns empty for an empty query', () => {
    const index = [entry('abc', [{ start: 0, end: 3, x: 0, y: 0, w: 30, h: 10 }])]
    expect(searchInIndex(index, '')).toEqual({ matches: [], capped: false, moreCount: 0 })
  })

  it('finds case-insensitive matches with interpolated rects', () => {
    const index = [entry('Hello World', [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }])]
    const { matches } = searchInIndex(index, 'WORLD')
    expect(matches).toHaveLength(1)
    expect(matches[0]!.pageIndex).toBe(0)
    // 'World' spans chars 6..11 of 11 -> x from 60 to 110
    expect(matches[0]!.rects).toHaveLength(1)
    const [x1, y1, x2, y2] = matches[0]!.rects[0]!
    expect(x1).toBeCloseTo(60)
    expect(y1).toBe(700)
    expect(x2).toBeCloseTo(110)
    expect(y2).toBe(712)
  })

  it('spans multiple items with one rect per item', () => {
    const index = [
      entry('abcdef', [
        { start: 0, end: 3, x: 0, y: 0, w: 30, h: 10 },
        { start: 3, end: 6, x: 30, y: 0, w: 30, h: 10 },
      ]),
    ]
    const { matches } = searchInIndex(index, 'cd')
    expect(matches).toHaveLength(1)
    expect(matches[0]!.rects).toHaveLength(2)
    expect(matches[0]!.rects[0]).toEqual([20, 0, 30, 10])
    expect(matches[0]!.rects[1]).toEqual([30, 0, 40, 10])
  })

  it('reports every occurrence and the correct page index', () => {
    const index = [
      entry('nothing here', [{ start: 0, end: 12, x: 0, y: 0, w: 120, h: 10 }]),
      entry('foo bar foo', [{ start: 0, end: 11, x: 0, y: 0, w: 110, h: 10 }]),
    ]
    const { matches } = searchInIndex(index, 'foo')
    expect(matches).toHaveLength(2)
    expect(matches.every((m) => m.pageIndex === 1)).toBe(true)
  })

  it('skips matches falling in EOL-only gaps with no item coverage', () => {
    // '\n' at chars 5..6 belongs to no item -> no rects -> match dropped
    const index = [
      entry('hello\nworld', [
        { start: 0, end: 5, x: 0, y: 0, w: 50, h: 10 },
        { start: 6, end: 11, x: 0, y: -20, w: 50, h: 10 },
      ]),
    ]
    expect(searchInIndex(index, 'hello').matches).toHaveLength(1)
    expect(searchInIndex(index, 'world').matches).toHaveLength(1)
    // The match itself spans the newline; rects come from both surrounding items
    expect(searchInIndex(index, 'hello\nworld').matches[0]!.rects).toHaveLength(2)
  })

  it('returns the full set un-capped when hits stay under the cap', () => {
    const index = [entry('a b a b a', [{ start: 0, end: 9, x: 0, y: 0, w: 90, h: 10 }])]
    const result = searchInIndex(index, 'a')
    expect(result.matches).toHaveLength(3)
    expect(result.capped).toBe(false)
    expect(result.moreCount).toBe(0)
  })

  it('reports no cap when the document has exactly MAX_MATCHES hits', () => {
    // 1000 'a's separated by spaces: exactly 1000 occurrences of 'a'
    const text = Array.from({ length: MAX_MATCHES }, () => 'a').join(' ')
    const index = [entry(text, [{ start: 0, end: text.length, x: 0, y: 0, w: text.length, h: 10 }])]
    const result = searchInIndex(index, 'a')
    expect(result.matches).toHaveLength(MAX_MATCHES)
    expect(result.capped).toBe(false)
    expect(result.moreCount).toBe(0)
  })

  it('caps at MAX_MATCHES with an honest more-count (BUG-1731)', () => {
    // 2000 adjacent 'a's: 2000 single-char occurrences -> 1000 kept + 1000 more
    const text = 'a'.repeat(2000)
    const index = [entry(text, [{ start: 0, end: 2000, x: 0, y: 0, w: 2000, h: 10 }])]
    const result = searchInIndex(index, 'a')
    expect(result.matches).toHaveLength(1000)
    expect(result.capped).toBe(true)
    expect(result.moreCount).toBe(1000)
  })

  it('counts hits past the cap across later pages, keeping first-page matches first', () => {
    // 600 hits on page 0, 700 on page 1 -> 1000 kept in document order, 300 more
    const page = (n: number): SearchIndex[number] => {
      const text = 'a '.repeat(n).trim()
      return entry(text, [{ start: 0, end: text.length, x: 0, y: 0, w: text.length, h: 10 }])
    }
    const index: SearchIndex = [page(600), page(700)]
    const result = searchInIndex(index, 'a')
    expect(result.matches).toHaveLength(1000)
    expect(result.matches[0]!.pageIndex).toBe(0)
    expect(result.matches[599]!.pageIndex).toBe(0)
    expect(result.matches[600]!.pageIndex).toBe(1)
    expect(result.capped).toBe(true)
    expect(result.moreCount).toBe(300)
  })

  describe('nextMatchIndex (wrap-around navigation over the capped set)', () => {
    it('wraps forward from the last match to the first', () => {
      expect(nextMatchIndex(MAX_MATCHES - 1, 1, MAX_MATCHES)).toBe(0)
    })

    it('wraps backward from the first match to the last', () => {
      expect(nextMatchIndex(0, -1, MAX_MATCHES)).toBe(MAX_MATCHES - 1)
    })

    it('steps normally inside the set and tolerates count 1', () => {
      expect(nextMatchIndex(0, 1, MAX_MATCHES)).toBe(1)
      expect(nextMatchIndex(5, -1, MAX_MATCHES)).toBe(4)
      expect(nextMatchIndex(0, -1, 1)).toBe(0)
      expect(nextMatchIndex(0, 1, 1)).toBe(0)
    })
  })
})

import { describe, expect, it } from 'vitest'
import { compareParagraphs, diffParagraphs, summarize } from '../src/renderer/editor/compare'

describe('compareParagraphs', () => {
  it('reports identical documents as all same', () => {
    const entries = compareParagraphs(['a', 'b'], ['a', 'b'])
    expect(entries.every((e) => e.kind === 'same')).toBe(true)
    expect(summarize(entries)).toEqual({ added: 0, removed: 0, changed: 0 })
  })

  it('detects an added paragraph', () => {
    const entries = compareParagraphs(['a', 'c'], ['a', 'b', 'c'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'added', 'same'])
    expect(entries[1].right).toBe('b')
  })

  it('detects a removed paragraph', () => {
    const entries = compareParagraphs(['a', 'b', 'c'], ['a', 'c'])
    expect(entries.map((e) => e.kind)).toEqual(['same', 'removed', 'same'])
    expect(entries[1].left).toBe('b')
  })

  it('merges adjacent remove+add into changed', () => {
    const entries = compareParagraphs(
      ['title', 'old content', 'ending'],
      ['title', 'new content', 'ending'],
    )
    expect(entries.map((e) => e.kind)).toEqual(['same', 'changed', 'same'])
    expect(entries[1]).toMatchObject({ left: 'old content', right: 'new content' })
  })

  it('handles empty documents', () => {
    expect(compareParagraphs([], [])).toEqual([])
    expect(compareParagraphs([], ['x'])[0].kind).toBe('added')
    expect(compareParagraphs(['x'], [])[0].kind).toBe('removed')
  })
})

describe('diffParagraphs cell budget (BUG-913: book-scale documents)', () => {
  const range = (n: number, tag: string) => Array.from({ length: n }, (_, i) => `${tag} ${i}`)

  it('uses the exact LCS while the matrix stays within the budget', () => {
    // 2000x2000 cells = exactly PARA_DIFF_BUDGET (inclusive bound)
    const n = 1999
    const paras = range(n, 'p')
    const { entries, degraded } = diffParagraphs(paras, paras.slice())
    expect(degraded).toBe(false)
    expect(entries.every((e) => e.kind === 'same')).toBe(true)
    // alignment quality check: LCS still skips a deleted middle paragraph
    const shifted = diffParagraphs(['a', 'gone', 'b'], ['a', 'b'])
    expect(shifted.degraded).toBe(false)
    expect(shifted.entries.map((e) => e.kind)).toEqual(['same', 'removed', 'same'])
  })

  it('degrades to index pairing above the budget', () => {
    // (2000+1)^2 = 4,004,001 cells > PARA_DIFF_BUDGET
    const { entries, degraded } = diffParagraphs(range(2000, 'left'), range(2000, 'right'))
    expect(degraded).toBe(true)
    expect(entries).toHaveLength(2000)
    expect(entries.every((e) => e.kind === 'changed')).toBe(true)
  })

  it('pairs length mismatches as changed + removed/added leftovers in degraded mode', () => {
    const { entries, degraded } = diffParagraphs(range(2500, 'l'), range(2000, 'r'))
    expect(degraded).toBe(true)
    const kinds = entries.map((e) => e.kind)
    expect(kinds.filter((k) => k === 'changed')).toHaveLength(2000)
    expect(kinds.filter((k) => k === 'removed')).toHaveLength(500)
  })

  it('finishes fast on identical book-scale documents (10k paragraphs)', () => {
    const n = 10_000
    const paras = range(n, 'paragraph')
    const start = Date.now()
    const { entries, degraded } = diffParagraphs(paras, paras.slice())
    const elapsed = Date.now() - start
    expect(degraded).toBe(true)
    expect(entries).toHaveLength(n)
    expect(entries.every((e) => e.kind === 'same')).toBe(true)
    // the fallback is O(n+m); an accidental revert to the O(n*m) LCS
    // (100M cells) takes tens of seconds — a generous ceiling catches that
    // without flaking on slow CI machines
    expect(elapsed).toBeLessThan(2000)
  })
})

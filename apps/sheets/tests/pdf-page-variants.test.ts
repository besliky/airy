import { describe, expect, it } from 'vitest'

import {
  evenPageRanges,
  pageRangesString,
  sheetPassesPlan,
  stitchPlan,
  variantForPage,
} from '../src/main/pdf-page-variants'

describe('variantForPage', () => {
  it('uses the odd templates everywhere when no variant is active', () => {
    const flags = { hasFirst: false, hasEven: false }
    expect([1, 2, 3, 4].map((page) => variantForPage(page, flags))).toEqual([
      'odd',
      'odd',
      'odd',
      'odd',
    ])
  })

  it('gives page 1 the first-page templates under differentFirst', () => {
    const flags = { hasFirst: true, hasEven: false }
    expect([1, 2, 3].map((page) => variantForPage(page, flags))).toEqual(['first', 'odd', 'odd'])
  })

  it('alternates odd/even under differentOddEven', () => {
    const flags = { hasFirst: false, hasEven: true }
    expect([1, 2, 3, 4, 5].map((page) => variantForPage(page, flags))).toEqual([
      'odd',
      'even',
      'odd',
      'even',
      'odd',
    ])
  })

  it('lets the first page win over parity when both flags are set', () => {
    const flags = { hasFirst: true, hasEven: true }
    expect([1, 2, 3, 4].map((page) => variantForPage(page, flags))).toEqual([
      'first',
      'even',
      'odd',
      'even',
    ])
  })
})

describe('evenPageRanges', () => {
  it('lists the even pages for Chromium pageRanges', () => {
    expect(evenPageRanges(1)).toBe('')
    expect(evenPageRanges(2)).toBe('2')
    expect(evenPageRanges(7)).toBe('2,4,6')
    expect(evenPageRanges(8)).toBe('2,4,6,8')
  })
})

describe('stitchPlan', () => {
  it('indexes each page into the pass that printed it', () => {
    expect(stitchPlan(5, { hasFirst: true, hasEven: true })).toEqual([
      { page: 1, source: 'first', index: 0 },
      { page: 2, source: 'even', index: 0 },
      { page: 3, source: 'odd', index: 2 },
      { page: 4, source: 'even', index: 1 },
      { page: 5, source: 'odd', index: 4 },
    ])
  })

  it('is the identity over the odd pass without variants', () => {
    expect(stitchPlan(3, { hasFirst: false, hasEven: false })).toEqual([
      { page: 1, source: 'odd', index: 0 },
      { page: 2, source: 'odd', index: 1 },
      { page: 3, source: 'odd', index: 2 },
    ])
    expect(stitchPlan(0, { hasFirst: true, hasEven: true })).toEqual([])
  })
})

describe('sheetPassesPlan', () => {
  const none = { hasFirst: false, hasEven: false }

  it('prints one pass per sheet without variants', () => {
    const plan = sheetPassesPlan(5, [2, 3], none)
    expect(plan.passes).toEqual([
      { sheet: 0, variant: 'odd', pages: [1, 2] },
      { sheet: 1, variant: 'odd', pages: [3, 4, 5] },
    ])
    expect(plan.steps).toEqual([
      { pass: 0, index: 0 },
      { pass: 0, index: 1 },
      { pass: 1, index: 0 },
      { pass: 1, index: 1 },
      { pass: 1, index: 2 },
    ])
  })

  it('splits every sheet by page variant under differentOddEven', () => {
    const plan = sheetPassesPlan(5, [2, 3], { hasFirst: false, hasEven: true })
    expect(plan.passes).toEqual([
      { sheet: 0, variant: 'odd', pages: [1] },
      { sheet: 0, variant: 'even', pages: [2] },
      { sheet: 1, variant: 'odd', pages: [3, 5] },
      { sheet: 1, variant: 'even', pages: [4] },
    ])
    expect(plan.steps).toEqual([
      { pass: 0, index: 0 },
      { pass: 1, index: 0 },
      { pass: 2, index: 0 },
      { pass: 3, index: 0 },
      { pass: 2, index: 1 },
    ])
  })

  it('gives page 1 the first-page pass of its owning sheet', () => {
    const plan = sheetPassesPlan(3, [2, 1], { hasFirst: true, hasEven: false })
    expect(plan.passes).toEqual([
      { sheet: 0, variant: 'first', pages: [1] },
      { sheet: 0, variant: 'odd', pages: [2] },
      { sheet: 1, variant: 'odd', pages: [3] },
    ])
    expect(plan.steps).toEqual([
      { pass: 0, index: 0 },
      { pass: 1, index: 0 },
      { pass: 2, index: 0 },
    ])
  })

  it('lets the last sheet own pages past the counted total (pagination drift)', () => {
    const plan = sheetPassesPlan(6, [2, 3], none)
    expect(plan.passes).toEqual([
      { sheet: 0, variant: 'odd', pages: [1, 2] },
      { sheet: 1, variant: 'odd', pages: [3, 4, 5, 6] },
    ])
    expect(plan.steps[5]).toEqual({ pass: 1, index: 3 })
  })

  it('skips zero-count sheets and handles an empty plan', () => {
    const plan = sheetPassesPlan(2, [0, 2], none)
    expect(plan.passes).toEqual([{ sheet: 1, variant: 'odd', pages: [1, 2] }])
    expect(sheetPassesPlan(0, [2, 3], none)).toEqual({ passes: [], steps: [] })
  })
})

describe('pageRangesString', () => {
  it('folds consecutive runs into first-last', () => {
    expect(pageRangesString([])).toBe('')
    expect(pageRangesString([3])).toBe('3')
    expect(pageRangesString([1, 2, 3])).toBe('1-3')
    expect(pageRangesString([1, 2, 3, 7, 9, 10, 11, 12])).toBe('1-3,7,9-12')
  })
})

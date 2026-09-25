/**
 * BUG-1718 regression: circular references must be statically detectable so
 * the editor can badge them — the engine resolves a cycle (A1=B1+1 /
 * B1=A1+1) in a single pass and silently shows the one-pass numbers, even
 * when the file asks for iterative calculation. Covers straight, self,
 * cross-sheet and range-level cycles, plus the negative space that must
 * never flag (chains, diamonds, out-of-range cells) and the badge text.
 */
import { describe, expect, it } from 'vitest'

import {
  circularRefAddress,
  findCircularFormulas,
  formatCircularRefList,
} from '../src/renderer/circular-refs'
import type { ClosureSheetInput } from '../src/renderer/formula-closure'

const sheet = (
  id: string,
  name: string,
  formulas: Record<string, string>,
  rowCount = 200,
  columnCount = 20,
): ClosureSheetInput => ({
  id,
  name,
  rowCount,
  columnCount,
  formulas: Object.entries(formulas).map(([address, formula]) => {
    const letters = /^([A-Z]+)([0-9]+)$/.exec(address)!
    let column = 0
    for (const character of letters[1] ?? '') column = column * 26 + character.charCodeAt(0) - 64
    return { row: Number(letters[2]) - 1, column: column - 1, formula }
  }),
})

const addresses = (sheets: ClosureSheetInput[], id: string) =>
  findCircularFormulas(sheets)
    .filter((hit) => hit.sheetId === id)
    .map((hit) => circularRefAddress(hit))

describe('findCircularFormulas', () => {
  it('finds the audited two-cell cycle and not its dependents', () => {
    const sheets = [
      sheet('s', 'Sheet1', {
        A1: '=B1+1',
        B1: '=A1+1',
        C1: '=A1',
        D1: '=SUM(A1:B1)',
      }),
    ]
    expect(addresses(sheets, 's')).toEqual(['A1', 'B1'])
  })

  it('finds a self-reference', () => {
    const sheets = [sheet('s', 'Sheet1', { A1: '=A1+1', B1: '=A1' })]
    expect(addresses(sheets, 's')).toEqual(['A1'])
  })

  it('finds a cycle through a whole-column reference', () => {
    const sheets = [sheet('s', 'Sheet1', { A1: '=SUM(B:B)', B1: '=A1+1' })]
    expect(addresses(sheets, 's')).toEqual(['A1', 'B1'])
  })

  it('finds cycles across sheets', () => {
    const sheets = [
      sheet('one', 'First', { A1: '=Second!A1+1' }),
      sheet('two', 'Second', { A1: '=First!A1+1', B1: '=B2' }),
    ]
    expect(addresses(sheets, 'one')).toEqual(['A1'])
    expect(addresses(sheets, 'two')).toEqual(['A1'])
  })

  it('finds a cycle through a referenced range', () => {
    const sheets = [sheet('s', 'Sheet1', { A1: '=SUM(B1:B10)', B5: '=A1*2' })]
    expect(addresses(sheets, 's')).toEqual(['A1', 'B5'])
  })

  it('does not flag plain chains or dependents outside the loop', () => {
    const sheets = [
      sheet('s', 'Sheet1', {
        A1: '=1',
        B1: '=A1+1',
        C1: '=B1*2',
        D1: '=SUM(B1:C1)',
      }),
    ]
    expect(findCircularFormulas(sheets)).toEqual([])
  })

  it('does not flag formulas merely referencing a range outside themselves', () => {
    const sheets = [sheet('s', 'Sheet1', { A1: '=SUM(B1:B10)', B20: '=A1*2' })]
    expect(findCircularFormulas(sheets)).toEqual([])
  })

  it('skips references to unknown sheets without crashing', () => {
    const sheets = [sheet('s', 'Sheet1', { A1: '=Missing!A1+1', B1: '=B2', B2: '=B1' })]
    expect(addresses(sheets, 's')).toEqual(['B1', 'B2'])
  })
})

describe('formatCircularRefList', () => {
  const hits = ['A1', 'B1', 'C1', 'D1', 'E1'].map((_, index) => ({
    sheetId: 's',
    row: 0,
    column: index,
  }))

  it('lists up to three addresses', () => {
    expect(formatCircularRefList(hits.slice(0, 2))).toBe('A1, B1')
    expect(formatCircularRefList(hits.slice(0, 3))).toBe('A1, B1, C1')
  })

  it('caps the list with a remaining count', () => {
    expect(formatCircularRefList(hits)).toBe('A1, B1, C1 +2')
  })
})

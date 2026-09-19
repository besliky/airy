// Unit tests for the shared case-insensitive matching primitive (BUG-1101).
// String.prototype.toLowerCase is not length-preserving: U+0130 (İ, Turkish
// dotted capital I) folds to "i" + U+0307, one code point becoming two code
// units, so every index after an İ in the lowered string is shifted against
// the original. findReplace over lowered indices used to silently eat or
// duplicate characters there; these tests pin the audit reproductions for
// all three consumers (markdown/html line ops, docx run ops).
import { describe, expect, it } from 'vitest'

import { caseInsensitiveRanges, replaceCaseInsensitive } from '../src/case-fold.js'

describe('caseInsensitiveRanges', () => {
  it('maps matches back through a length-changing fold (İ before the hit)', () => {
    // folded: "i̇stanbul kelime not" — the match at folded index 10 must
    // address original index 9, not 10
    expect(caseInsensitiveRanges('İstanbul kelime not', 'kelime')).toEqual([{ start: 9, end: 15 }])
    // the audit's minimal isolation: 'b' sits at original index 2
    expect(caseInsensitiveRanges('İabc', 'b')).toEqual([{ start: 2, end: 3 }])
  })

  it('maps every unit of an expanding code point to its source index', () => {
    // matching "i" against İ consumes the whole code point (a partial one
    // cannot be represented); matching the full two-unit fold does too
    expect(caseInsensitiveRanges('İabc', 'i')).toEqual([{ start: 0, end: 1 }])
    expect(caseInsensitiveRanges('aİb', '\u0069\u0307')).toEqual([{ start: 1, end: 2 }])
    expect(caseInsensitiveRanges('İi', 'i')).toEqual([
      { start: 0, end: 1 },
      { start: 1, end: 2 },
    ])
  })

  it('keeps surrogate pairs whole (astral code points)', () => {
    expect(caseInsensitiveRanges('\u{1F600}İx', 'x')).toEqual([{ start: 3, end: 4 }])
    expect(caseInsensitiveRanges('\u{1D400}bc', '\u{1D400}')).toEqual([{ start: 0, end: 2 }])
  })

  it('finds non-overlapping repeats and nothing for an empty needle', () => {
    expect(caseInsensitiveRanges('İaİa', 'a')).toEqual([
      { start: 1, end: 2 },
      { start: 3, end: 4 },
    ])
    expect(caseInsensitiveRanges('anything', '')).toEqual([])
  })
})

describe('replaceCaseInsensitive', () => {
  it('replaces after an İ without eating the tail (audit live repro)', () => {
    // the line-session repro: "İstanbul kWORDnot" before the fix
    const outcome = replaceCaseInsensitive('İstanbul kelime not', 'kelime', 'WORD')
    expect(outcome.text).toBe('İstanbul WORD not')
    expect(outcome.count).toBe(1)
  })

  it('passes the audit isolation that lost the last character', () => {
    // "İabX" before the fix: the c after the match was sliced away
    expect(replaceCaseInsensitive('İabc', 'b', 'X')).toEqual({ text: 'İaXc', count: 1 })
  })

  it('leaves text without matches untouched (and counts zero)', () => {
    expect(replaceCaseInsensitive('İstanbul', 'kelime', 'X')).toEqual({
      text: 'İstanbul',
      count: 0,
    })
  })

  it('replaces several matches around İ occurrences', () => {
    expect(replaceCaseInsensitive('İxİy x', 'x', 'X')).toEqual({ text: 'İXİy X', count: 2 })
    // a needle that IS the İ fold replaces the whole code point
    expect(replaceCaseInsensitive('aİb', 'i', 'Y')).toEqual({ text: 'aYb', count: 1 })
  })

  it('keeps plain ASCII case-insensitive behavior', () => {
    expect(replaceCaseInsensitive('Revenue grew. revenue fell.', 'revenue', 'sales')).toEqual({
      text: 'sales grew. sales fell.',
      count: 2,
    })
  })

  it('preserves whole-string fold semantics (Greek final sigma stays matchable)', () => {
    // "ΟΛΟΣ".toLowerCase() is "ολος" (final sigma), while per-code-point Σ
    // folds to σ — the search must run on the whole-string fold, exactly
    // like the pre-fix code, or word-final sigma would stop matching
    expect(replaceCaseInsensitive('ΟΛΟΣ', 'ος', 'OS')).toEqual({ text: 'ΟΛOS', count: 1 })
    expect(replaceCaseInsensitive('ΟΛΟΣ', 'ολος', 'ALL')).toEqual({ text: 'ALL', count: 1 })
  })
})

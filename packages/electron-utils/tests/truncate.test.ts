import { describe, expect, it } from 'vitest'

import { truncateByCodePoints } from '../src/truncate'

describe('truncateByCodePoints', () => {
  it('keeps short strings untouched', () => {
    expect(truncateByCodePoints('report.docx', 80)).toBe('report.docx')
    expect(truncateByCodePoints('', 80)).toBe('')
  })

  it('cuts a too-long ASCII string exactly like slice would', () => {
    expect(truncateByCodePoints('a'.repeat(90), 80)).toBe('a'.repeat(80))
  })

  it('never splits a surrogate pair at the cap (BUG-412)', () => {
    // 40 astral code points = 80 UTF-16 code units; the legacy slice(0, 79)
    // cut the last pair in half, leaving a lone high surrogate that cannot be
    // encoded to UTF-8 — the filesystem then rejects the whole name
    const name = '🦄'.repeat(40)
    const legacy = name.slice(0, 79)
    expect(legacy).toMatch(/[\ud800-\udbff](?![\udc00-\udfff])/) // broken pair
    // 40 pairs + '.docx' = 45 code points: the whole suffix fits under the cap
    const capped = truncateByCodePoints(name + '.docx', 80)
    expect(capped).toBe(name + '.docx')
    expect(capped).toMatch(/^(?:🦄)+\.docx$/) // every code unit belongs to a whole pair
    // a pure-astral string caps at the code-point count (80 pairs = 160 units)
    expect(truncateByCodePoints('🦄'.repeat(85), 80)).toBe('🦄'.repeat(80))
  })

  it('caps BMP-heavy and astral-heavy strings by code-point count', () => {
    expect(truncateByCodePoints('ab😀cd', 3)).toBe('ab😀')
    expect(truncateByCodePoints('абвгд', 4)).toBe('абвг')
  })

  it('rejects a non-positive or fractional cap', () => {
    expect(() => truncateByCodePoints('x', 0)).toThrow(RangeError)
    expect(() => truncateByCodePoints('x', 1.5)).toThrow(RangeError)
  })
})

import { describe, expect, it } from 'vitest'

import { HOME_PATHS_CAP, stringPathsCapped } from '../src/shared/home-paths'

describe('stringPathsCapped', () => {
  it('keeps the string entries in order and drops junk', () => {
    expect(stringPathsCapped(['/a', 42, '/b', null, '  '])).toEqual(['/a', '/b', '  '])
    expect(stringPathsCapped('/a')).toEqual([])
    expect(stringPathsCapped(null)).toEqual([])
  })

  it('caps the accepted list and ignores extras', () => {
    const many = Array.from({ length: HOME_PATHS_CAP + 1_000 }, (_, i) => `/tmp/f${i}`)
    const capped = stringPathsCapped(many)
    expect(capped.length).toBe(HOME_PATHS_CAP)
    expect(capped[0]).toBe('/tmp/f0')
    expect(capped[HOME_PATHS_CAP - 1]).toBe(`/tmp/f${HOME_PATHS_CAP - 1}`)
  })

  it('covers full project catalogs at production level (BUG-1676: was 256)', () => {
    expect(HOME_PATHS_CAP).toBeGreaterThanOrEqual(1_000)
    // a 300-file project used to be silently truncated to 256 entries
    const catalog = Array.from({ length: 300 }, (_, i) => `/proj/p${String(i).padStart(4, '0')}`)
    expect(stringPathsCapped(catalog)).toHaveLength(300)
    expect(stringPathsCapped(catalog)).toContain('/proj/p0280')
  })

  it('honors an explicit cap and rejects nonsensical ones', () => {
    expect(stringPathsCapped(['/a', '/b', '/c'], 2)).toEqual(['/a', '/b'])
    expect(stringPathsCapped(['/a'], 0)).toEqual([])
    expect(stringPathsCapped(['/a'], -5)).toEqual([])
  })

  it('counts only string entries against the cap', () => {
    const junk = Array.from({ length: 50 }, () => 42)
    const capped = stringPathsCapped([...junk, '/a', '/b', '/c'], 2)
    expect(capped).toEqual(['/a', '/b'])
  })
})

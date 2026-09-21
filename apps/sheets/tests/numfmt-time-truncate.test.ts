/**
 * BUG-1525 (audited as BUG-1503) time display truncation: Excel floors a
 * time to the least significant unit the pattern shows — h:mm renders
 * 14:29:59.997 as 14:29 — while numfmt rounds and printed 14:30. The fix
 * truncates the serial down to the displayed unit before formatting, on
 * the shared display path the grid and the print layout both read.
 */
import { describe, expect, it } from 'vitest'

import { numfmt } from '@univerjs/core'

import { fixFormattedValue, truncateTimeForPattern } from '../src/renderer/numfmt-fix'

/// 14:29:59.997 as an Excel day serial.
const ALMOST_HALF = 0.604166666

describe('truncateTimeForPattern', () => {
  it('floors to the minute for h:mm (the audit case)', () => {
    expect(truncateTimeForPattern('h:mm', ALMOST_HALF)).toBe(Math.floor(ALMOST_HALF * 1440) / 1440)
    expect(numfmt.format('h:mm', truncateTimeForPattern('h:mm', ALMOST_HALF)!)).toBe('14:29')
  })

  it('floors to the second for h:mm:ss and to milliseconds for s.000', () => {
    const second = truncateTimeForPattern('h:mm:ss', ALMOST_HALF)!
    expect(second).toBe(Math.floor(ALMOST_HALF * 86400) / 86400)
    expect(numfmt.format('h:mm:ss', second)).toBe('14:29:59')
    const milli = truncateTimeForPattern('s.000', ALMOST_HALF)!
    expect(milli).toBe(Math.floor(ALMOST_HALF * 86400000) / 86400000)
    expect(numfmt.format('s.000', milli)).toBe('59.999')
  })

  it('floors to the hour for an hours-only pattern and minutes for elapsed [h]:mm', () => {
    const hour = truncateTimeForPattern('h', ALMOST_HALF)!
    expect(numfmt.format('h', hour)).toBe('14')
    expect(truncateTimeForPattern('[h]:mm', ALMOST_HALF)).toBe(
      Math.floor(ALMOST_HALF * 1440) / 1440,
    )
  })

  it('keeps date+time patterns truncated to their time unit, not to the day', () => {
    const serial = 45291 + ALMOST_HALF
    expect(truncateTimeForPattern('m/d/yy h:mm', serial)).toBe(Math.floor(serial * 1440) / 1440)
    expect(numfmt.format('m/d/yy h:mm', truncateTimeForPattern('m/d/yy h:mm', serial)!)).toBe(
      '12/31/23 14:29',
    )
  })

  it('ignores meridiem tokens, quoted literals, colors and locale tags', () => {
    // AM/PM carries an m that is not a minute; "ms" is a quoted literal;
    // [Red] and [$-409] are opaque bracket blocks.
    expect(truncateTimeForPattern('h:mm AM/PM', ALMOST_HALF)).toBe(
      Math.floor(ALMOST_HALF * 1440) / 1440,
    )
    expect(truncateTimeForPattern('"ms" h:mm', ALMOST_HALF)).toBe(
      Math.floor(ALMOST_HALF * 1440) / 1440,
    )
    expect(truncateTimeForPattern('[Red][$-409]h:mm', ALMOST_HALF)).toBe(
      Math.floor(ALMOST_HALF * 1440) / 1440,
    )
  })

  it('returns null for date-only patterns, whole values and negative serials', () => {
    expect(truncateTimeForPattern('m/d/yy', ALMOST_HALF)).toBeNull()
    expect(truncateTimeForPattern('yyyy"ms"', 45291)).toBeNull()
    expect(truncateTimeForPattern('h:mm', 0.5)).toBeNull()
    expect(truncateTimeForPattern('h:mm', -ALMOST_HALF)).toBeNull()
    expect(truncateTimeForPattern('h:mm', Number.NaN)).toBeNull()
  })
})

describe('fixFormattedValue time truncation', () => {
  it('corrects the rounded display numfmt left behind', () => {
    // numfmt's own render rounds up to 14:30; the fix returns 14:29.
    expect(fixFormattedValue('h:mm', ALMOST_HALF, '14:30')).toBe('14:29')
    expect(fixFormattedValue('h:mm:ss', ALMOST_HALF, '14:30:00')).toBe('14:29:59')
  })

  it('leaves already-truncated displays untouched', () => {
    expect(fixFormattedValue('h:mm', Math.floor(ALMOST_HALF * 1440) / 1440, '14:29')).toBeNull()
  })

  it('does not touch number patterns (the decimal round still applies)', () => {
    expect(fixFormattedValue('0.00', 1.005, '1.00')).toBe('1.01')
  })
})

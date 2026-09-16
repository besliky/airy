import { describe, expect, it } from 'vitest'

/**
 * Tab-switch accelerators (src/main/tab-accelerators.ts): Ctrl/Cmd+1..8
 * select tab N (Home is tab 1), Ctrl/Cmd+9 the last tab. Digits the editors
 * reserve for their own Word/Excel shortcuts must never be captured.
 */
import {
  RESERVED_TAB_DIGITS,
  switchableDigitsForKind,
  switchDigitFromInput,
  tabIndexForDigit,
} from '../src/main/tab-accelerators'

describe('tabIndexForDigit', () => {
  it('maps 1..8 to their strip index and 9 to the last tab', () => {
    expect(tabIndexForDigit(1, 5)).toBe(0)
    expect(tabIndexForDigit(8, 9)).toBe(7)
    expect(tabIndexForDigit(9, 5)).toBe(4)
  })

  it('returns null beyond the strip or for invalid digits', () => {
    expect(tabIndexForDigit(3, 2)).toBeNull()
    expect(tabIndexForDigit(1, 0)).toBeNull()
    expect(tabIndexForDigit(0, 5)).toBeNull()
    expect(tabIndexForDigit(10, 5)).toBeNull()
    expect(tabIndexForDigit(1.5, 5)).toBeNull()
  })
})

describe('switchDigitFromInput', () => {
  const input = (over: Partial<Parameters<typeof switchDigitFromInput>[0]>) => ({
    type: 'keyDown',
    control: true,
    meta: false,
    alt: false,
    shift: false,
    code: 'Digit3',
    ...over,
  })

  it('accepts Ctrl+digit and Cmd+digit', () => {
    expect(switchDigitFromInput(input({}))).toBe(3)
    expect(switchDigitFromInput(input({ control: false, meta: true }))).toBe(3)
  })

  it('ignores keyUp, plain digits, and other keys', () => {
    expect(switchDigitFromInput(input({ type: 'keyUp' }))).toBeNull()
    expect(switchDigitFromInput(input({ control: false }))).toBeNull()
    expect(switchDigitFromInput(input({ code: 'KeyA' }))).toBeNull()
    expect(switchDigitFromInput(input({ code: 'Digit0' }))).toBeNull()
  })

  it('ignores modifier combinations the editors own', () => {
    // sheets Ctrl+Shift+9 unhide, docs Alt+Cmd+1..3 heading styles
    expect(switchDigitFromInput(input({ shift: true }))).toBeNull()
    expect(switchDigitFromInput(input({ alt: true }))).toBeNull()
  })
})

describe('reserved digits', () => {
  it('docs keeps Word line spacing and marks; sheets keeps Excel hide rows', () => {
    expect([...RESERVED_TAB_DIGITS.docs!].sort()).toEqual([1, 2, 5, 8])
    expect([...RESERVED_TAB_DIGITS.sheets!]).toEqual([9])
  })

  it('other kinds reserve nothing', () => {
    for (const kind of ['home', 'slides', 'pdf', 'markdown', 'html'] as const) {
      expect(RESERVED_TAB_DIGITS[kind]).toBeUndefined()
    }
  })

  it('switchable digits exclude the reserved ones', () => {
    expect(switchableDigitsForKind('docs')).toEqual([3, 4, 6, 7, 9])
    expect(switchableDigitsForKind('sheets')).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
    expect(switchableDigitsForKind('home')).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9])
  })
})

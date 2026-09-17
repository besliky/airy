import { describe, expect, it } from 'vitest'

import { isHomeSender } from '../src/main/home-sender-guard'

describe('isHomeSender', () => {
  it('accepts only the captured Home tab webContents id', () => {
    expect(isHomeSender(12, 12)).toBe(true)
    expect(isHomeSender(12, 13)).toBe(false)
    expect(isHomeSender(0, 0)).toBe(true)
  })

  it('rejects when no shell window exists or the sender id is missing', () => {
    expect(isHomeSender(null, 12)).toBe(false)
    expect(isHomeSender(undefined, 12)).toBe(false)
    expect(isHomeSender(12, undefined)).toBe(false)
    expect(isHomeSender(null, undefined)).toBe(false)
  })

  it('does not confuse falsy-looking ids (0 is a real webContents id)', () => {
    expect(isHomeSender(0, 0)).toBe(true)
    expect(isHomeSender(0, 1)).toBe(false)
  })
})

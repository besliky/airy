import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { beforeEach, describe, expect, it, vi } from 'vitest'

// vi.mock factories are hoisted above the const declarations, so the mock
// fn must be created via vi.hoisted to stay reachable from the factory.
const setAccessibilitySupportEnabled = vi.hoisted(() => vi.fn())

vi.mock('electron', () => ({
  app: { setAccessibilitySupportEnabled },
}))

import { applyAccessibilityPolicy } from '../src/main/accessibility-policy'

/**
 * Lazy accessibility policy for the sheets standalone entry (PERF-1727b),
 * same as the shell (PERF-1700c): support must stay off until Chromium's own
 * assistive-tech detection turns it on, and only the explicit
 * AIRY_FORCE_A11Y=1 escape hatch may force it.
 */

describe('applyAccessibilityPolicy (sheets standalone)', () => {
  const saved = process.env.AIRY_FORCE_A11Y

  beforeEach(() => {
    setAccessibilitySupportEnabled.mockClear()
    if (saved === undefined) delete process.env.AIRY_FORCE_A11Y
    else process.env.AIRY_FORCE_A11Y = saved
  })

  it('leaves support to Chromium detection when no override is set', () => {
    delete process.env.AIRY_FORCE_A11Y
    expect(applyAccessibilityPolicy()).toBe('detection')
    expect(setAccessibilitySupportEnabled).not.toHaveBeenCalled()
  })

  it('forces support on when AIRY_FORCE_A11Y=1', () => {
    process.env.AIRY_FORCE_A11Y = '1'
    expect(applyAccessibilityPolicy()).toBe('forced')
    expect(setAccessibilitySupportEnabled).toHaveBeenCalledWith(true)
  })

  it('treats any other AIRY_FORCE_A11Y value as no override', () => {
    process.env.AIRY_FORCE_A11Y = '0'
    expect(applyAccessibilityPolicy()).toBe('detection')
    expect(setAccessibilitySupportEnabled).not.toHaveBeenCalled()
  })

  it('is wired into the standalone entry instead of an unconditional force', () => {
    // Source contract: startSheetsStandalone must route through the policy
    // module and sheets-main must not force accessibility support directly.
    const src = readFileSync(join(__dirname, '../src/main/sheets-main.ts'), 'utf8')
    expect(src).toContain("from './accessibility-policy'")
    expect(src).not.toContain('setAccessibilitySupportEnabled')
  })
})

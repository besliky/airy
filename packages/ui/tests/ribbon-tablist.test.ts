import { describe, expect, it } from 'vitest'

import {
  nextRibbonTabIndex,
  ribbonPanelProps,
  ribbonTabId,
  wrapRibbonTabIndex,
} from '../src/ribbon-tablist'

describe('wrapRibbonTabIndex', () => {
  it('wraps at both ends', () => {
    expect(wrapRibbonTabIndex(0, 3, -1)).toBe(2)
    expect(wrapRibbonTabIndex(2, 3, 1)).toBe(0)
    expect(wrapRibbonTabIndex(0, 3, 1)).toBe(1)
    expect(wrapRibbonTabIndex(1, 3, -1)).toBe(0)
  })

  it('returns -1 without tabs', () => {
    expect(wrapRibbonTabIndex(0, 0, 1)).toBe(-1)
  })
})

describe('nextRibbonTabIndex', () => {
  const keys = ['Home', 'Insert', 'Draw'] as const

  it('moves right and wraps in LTR', () => {
    expect(nextRibbonTabIndex(0, keys.length, 'ArrowRight')).toBe(1)
    expect(nextRibbonTabIndex(2, keys.length, 'ArrowRight')).toBe(0)
  })

  it('mirrors the horizontal arrows in RTL', () => {
    expect(nextRibbonTabIndex(1, keys.length, 'ArrowRight', true)).toBe(0)
    expect(nextRibbonTabIndex(0, keys.length, 'ArrowLeft', true)).toBe(1)
  })

  it('jumps to the strip ends with Home and End', () => {
    expect(nextRibbonTabIndex(1, keys.length, 'Home')).toBe(0)
    expect(nextRibbonTabIndex(0, keys.length, 'End')).toBe(keys.length - 1)
  })

  it('ignores keys the tablist does not own', () => {
    for (const key of ['ArrowUp', 'ArrowDown', 'Enter', ' ', 'Tab', 'a', 'F1']) {
      expect(nextRibbonTabIndex(1, keys.length, key)).toBe(-1)
    }
  })

  it('returns -1 for an empty strip', () => {
    expect(nextRibbonTabIndex(0, 0, 'ArrowRight')).toBe(-1)
  })
})

describe('ribbonTabId', () => {
  it('slugs characters HTML ids must not contain', () => {
    expect(ribbonTabId('sheets-ribbon', 'home')).toBe('sheets-ribbon-tab-home')
    expect(ribbonTabId('sheets-ribbon', 'Page Layout')).toBe('sheets-ribbon-tab-Page-Layout')
    expect(ribbonTabId('sheets-ribbon', 'Chart Design')).toBe('sheets-ribbon-tab-Chart-Design')
  })
})

describe('ribbonPanelProps', () => {
  it('labels the panel by the active tab id', () => {
    expect(ribbonPanelProps('docs-ribbon', 'home')).toEqual({
      role: 'tabpanel',
      id: 'docs-ribbon-panel',
      'aria-labelledby': 'docs-ribbon-tab-home',
    })
  })

  it('omits aria-labelledby before any selection', () => {
    expect(ribbonPanelProps('docs-ribbon', null)).toEqual({
      role: 'tabpanel',
      id: 'docs-ribbon-panel',
      'aria-labelledby': undefined,
    })
  })
})

import { describe, expect, it } from 'vitest'
import {
  SHADOW_PRESETS,
  shadowCss,
  shadowDecls,
  shadowPresetKey,
} from '../src/renderer/editor/shadow-effects'

describe('shadowCss approximation', () => {
  it('converts the DrawingML direction/distance to CSS px offsets', () => {
    // 45° bottom-right: dx = dy = 38100 EMU * cos45 / 9525 ≈ 2.8px
    const css = shadowCss(SHADOW_PRESETS[3].shadow!)
    expect(css).toBe('filter:drop-shadow(2.8px 2.8px 5.3px rgba(0,0,0,0.43))')
  })

  it('maps the inner preset to an inset box-shadow', () => {
    const inner = SHADOW_PRESETS.find((p) => p.shadow?.inner)!.shadow!
    expect(shadowCss(inner)).toMatch(/^box-shadow:inset /)
  })

  it('returns null for invalid colors', () => {
    expect(shadowCss({ ...SHADOW_PRESETS[1].shadow!, color: 'XYZ' })).toBeNull()
  })

  it('exposes the approximation as style properties for the gallery preview', () => {
    expect(shadowDecls(SHADOW_PRESETS[1].shadow!)).toEqual({
      filter: 'drop-shadow(2.7px 0.0px 0.0px rgba(0,0,0,0.43))',
    })
  })
})

describe('shadowPresetKey', () => {
  it('identifies the gallery entry matching a shadow (and None for null)', () => {
    expect(shadowPresetKey(null)).toBe('ribbonShadowNone')
    expect(shadowPresetKey(SHADOW_PRESETS[3].shadow)).toBe('ribbonShadowOffsetDiagBottomRight')
    // a hand-tuned shadow that matches none of the presets
    expect(shadowPresetKey({ blurRadEmu: 1, distEmu: 2, dirEmu: 3, color: '123456' })).toBeNull()
  })
})

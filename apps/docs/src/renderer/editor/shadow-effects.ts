/**
 * Picture/shape shadow presets (Word's Picture Effects → Shadow gallery
 * slice: the outer-offset family plus one inner preset) and the CSS
 * approximation used to render an authored or imported a:effectLst shadow.
 *
 * Preset numbers follow the values Word itself writes for the gallery
 * entries (EMU; 43% ≈ Word's effectStyle alpha). Effects outside this slice
 * — glow, reflection, soft edges, 3-D — are intentionally not modeled; such
 * drawings keep their original bytes and simply render unshadowed.
 */
import type { ShadowEffect } from '@airy-office/docx-engine'

export interface ShadowPreset {
  /** i18n key of the gallery label */
  labelKey: string
  /** null = the "No Shadow" entry */
  shadow: ShadowEffect | null
}

const EMU_PER_PX = 9525

/** opaque black at Word's shadow opacity (~43%) unless a preset overrides it */
const outer = (
  dirEmu: number,
  distEmu: number,
  blurRadEmu: number,
  alphaPct = 43,
): ShadowEffect => ({ dirEmu, distEmu, blurRadEmu, color: '000000', alphaPct })

export const SHADOW_PRESETS: ShadowPreset[] = [
  { labelKey: 'ribbonShadowNone', shadow: null },
  { labelKey: 'ribbonShadowOffsetRight', shadow: outer(0, 25400, 0) },
  { labelKey: 'ribbonShadowOffsetBottom', shadow: outer(5400000, 25400, 0) },
  { labelKey: 'ribbonShadowOffsetDiagBottomRight', shadow: outer(2700000, 38100, 50800) },
  { labelKey: 'ribbonShadowOffsetCenter', shadow: outer(5400000, 0, 76200) },
  {
    labelKey: 'ribbonShadowInnerDiagTopLeft',
    shadow: { ...outer(2700000, 38100, 50800, 35), inner: true },
  },
]

/** match a shadow against the gallery to highlight the active entry */
export function shadowPresetKey(shadow: ShadowEffect | null | undefined): string | null {
  if (!shadow) return 'ribbonShadowNone'
  return (
    SHADOW_PRESETS.find(
      (p) =>
        p.shadow != null &&
        p.shadow.dirEmu === shadow.dirEmu &&
        p.shadow.distEmu === shadow.distEmu &&
        p.shadow.blurRadEmu === shadow.blurRadEmu &&
        p.shadow.inner === shadow.inner,
    )?.labelKey ?? null
  )
}

/**
 * CSS approximation of a DrawingML shadow, as style properties: outer
 * shadows use drop-shadow so the silhouette follows the real shape (rotated
 * pictures, transparent PNGs, SVG preset geometry); the inner preset
 * approximates as an inset box-shadow. Returns null for invalid models.
 */
export function shadowDecls(shadow: ShadowEffect): { filter?: string; boxShadow?: string } | null {
  if (!shadow.color || !/^[0-9A-Fa-f]{6}$/i.test(shadow.color)) return null
  const rad = ((shadow.dirEmu / 60000) * Math.PI) / 180
  const dx = (shadow.distEmu * Math.cos(rad)) / EMU_PER_PX
  const dy = (shadow.distEmu * Math.sin(rad)) / EMU_PER_PX
  const blur = shadow.blurRadEmu / EMU_PER_PX
  const alpha = Math.min(1, Math.max(0, (shadow.alphaPct ?? 100) / 100))
  const r = parseInt(shadow.color.slice(0, 2), 16)
  const g = parseInt(shadow.color.slice(2, 4), 16)
  const b = parseInt(shadow.color.slice(4, 6), 16)
  const color = `rgba(${r},${g},${b},${alpha.toFixed(2)})`
  const px = (v: number) => (Math.round(v * 10) / 10).toFixed(1)
  return shadow.inner
    ? { boxShadow: `inset ${px(dx)}px ${px(dy)}px ${px(blur)}px ${color}` }
    : { filter: `drop-shadow(${px(dx)}px ${px(dy)}px ${px(blur)}px ${color})` }
}

/** the same approximation as one full CSS declaration (property included) */
export function shadowCss(shadow: ShadowEffect | null | undefined): string | null {
  if (!shadow) return null
  const decls = shadowDecls(shadow)
  if (!decls) return null
  return decls.filter ? `filter:${decls.filter}` : `box-shadow:${decls.boxShadow}`
}

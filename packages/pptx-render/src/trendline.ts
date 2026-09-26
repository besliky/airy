/**
 * Linear trendline math (c:trendline, linear regression).
 *
 * Pure numeric helpers shared by the chart builders: the least-squares fit
 * (ordinary least squares over paired samples) and the equation / R² label
 * formatting PowerPoint shows next to the line.
 */

export interface LinearFit {
  /** Slope (y = slope·x + intercept) */
  slope: number
  intercept: number
  /** Coefficient of determination of the fit (0..1 for ordinary data) */
  r2: number
}

/**
 * Ordinary least squares y = slope·x + intercept over paired samples.
 * Returns null when a line is not determined (fewer than 2 pairs or a
 * degenerate x domain — every x equal, every y equal is fine: slope 0).
 */
export function linearFit(pts: Array<{ x: number; y: number }>): LinearFit | null {
  const n = pts.length
  if (n < 2) return null
  let sx = 0
  let sy = 0
  for (const p of pts) {
    sx += p.x
    sy += p.y
  }
  const mx = sx / n
  const my = sy / n
  let sxx = 0
  let sxy = 0
  for (const p of pts) {
    sxx += (p.x - mx) * (p.x - mx)
    sxy += (p.x - mx) * (p.y - my)
  }
  if (sxx <= 0) return null
  const slope = sxy / sxx
  const intercept = my - slope * mx
  // R² = SSR/SST; a perfect horizontal fit (SST = 0) counts as r² 1
  let sst = 0
  let ssr = 0
  for (const p of pts) {
    const f = slope * p.x + intercept
    sst += (p.y - my) * (p.y - my)
    ssr += (p.y - f) * (p.y - f)
  }
  const r2 = sst <= 0 ? 1 : Math.max(0, 1 - ssr / sst)
  return { slope, intercept, r2 }
}

/** Compact coefficient format (PowerPoint "General"-like: ≤3 decimals, no trailing zeros). */
function fmtCoef(v: number): string {
  const s = Math.abs(v) >= 1e-6 ? v.toFixed(3) : v.toExponential(2)
  return s.replace(/\.?0+$/, '')
}

/**
 * Trendline equation label ("y = 0.5x + 1.5"); PowerPoint drops coefficients of
 * 1, zero terms, and writes negative intercepts as "- …".
 */
export function formatTrendEquation(slope: number, intercept: number): string {
  const slopePart = slope === 1 ? '' : slope === -1 ? '-' : `${fmtCoef(slope)}`
  const interceptPart =
    intercept === 0 ? '' : intercept < 0 ? ` - ${fmtCoef(-intercept)}` : ` + ${fmtCoef(intercept)}`
  return `y = ${slopePart}x${interceptPart}`
}

/** R² label ("R² = 0.9986", PowerPoint's 4-decimal convention). */
export function formatRSquared(r2: number): string {
  return `R² = ${r2.toFixed(4)}`
}

/** Document zoom window (percent), matching Word's 50–400 limits. Every zoom
 *  input — wheel deltas, the page-fit computation, the ±10 step buttons and
 *  the slider bounds — goes through clampDocZoom so the entry points can
 *  never disagree on the window. Rounding stays with the callers: the fit
 *  path floors on purpose (rounding up would push the page past the pane
 *  edge), the wheel allows fractional per-tick deltas. */
export const MIN_DOC_ZOOM = 50
export const MAX_DOC_ZOOM = 400

/// Clamp a zoom percent to the window; non-finite input falls back to 100.
export function clampDocZoom(percent: number): number {
  if (!Number.isFinite(percent)) return 100
  return Math.min(MAX_DOC_ZOOM, Math.max(MIN_DOC_ZOOM, percent))
}

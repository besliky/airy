/**
 * Zoom-state coherence for the edit canvas (BUG-1726).
 *
 * A drag's screen→page delta conversion happens inside Konva from the CSS
 * transform that is actually applied to the stage, so "cursor and shape move
 * together" holds only while the applied scale matches what the rest of the
 * app believes. The zoom gestures (slider/pinch/buttons) apply their value to
 * the CSS transform immediately and commit the state debounced, which means
 * state/refs can trail the visible scale for ~150ms. Any layout ripple in that
 * window (classic scrollbars appearing, a pane toggling) used to make the
 * resize-observer re-fit decision from the STALE zoom — slamming the canvas
 * back to fit under an in-flight zoom (label ≠ canvas) and, mid-drag,
 * teleporting the dragged shape relative to the cursor (BUG-1726: model delta
 * overshot the cursor by canvasW/pageW; some gestures never started because
 * the shape moved out from under the pointer before mousedown landed).
 *
 * The two helpers below are the whole contract:
 *  - appliedStageZoom measures the scale the canvas is REALLY displaying;
 *  - resizeFitDecision decides the observer's re-fit/clamp from that value.
 */

/** The zoom the stage is actually displaying right now (CSS transform vs the
 * unscaled layout size); null when the element is not measurable yet. */
export function appliedStageZoom(
  el: {
    offsetWidth: number
    getBoundingClientRect(): { width: number }
  } | null,
): number | null {
  if (!el || !(el.offsetWidth > 0)) return null
  return el.getBoundingClientRect().width / el.offsetWidth
}

export interface FitDecisionInput {
  /** Scale the canvas is displaying right now (appliedStageZoom; null = not measurable yet) */
  appliedZoom: number | null
  /** Fallback when no applied value is measurable (pre-mount) */
  stateZoom: number
  /** Uncapped fit for the current viewport ("can it still fit?" test) */
  rawFitZoom: number
  /** Clamped fit — the value fit mode re-fits to */
  fitZoom: number
  /** Fit zoom captured when fit mode last applied; null = never fitted */
  lastFitZoom: number | null
  /** True when the wrap's BORDER box changed (a real window/pane resize; scrollbar ripples don't count) */
  outerResized: boolean
}

export type FitDecision = { action: 'refit'; zoom: number } | { action: 'none' }

/** Resize-observer decision: follow with a re-fit while in fit mode, and clamp
 * a manual zoom back to fit on a real container shrink. Decisions read the
 * APPLIED zoom, never the lagging state — a zoom gesture in its commit window
 * (or a manual zoom below fit) must never be read as "still in fit mode". */
export function resizeFitDecision(input: FitDecisionInput): FitDecision {
  const zoom = input.appliedZoom ?? input.stateZoom
  const inFitMode = input.lastFitZoom != null && Math.abs(zoom - input.lastFitZoom) <= 0.001
  if (!inFitMode) {
    // The overflow test uses the uncapped ratio: on large windows a manual
    // zoom above the 1.5 fit cap can still fit and must not be wiped
    if (zoom <= input.rawFitZoom + 0.001) return { action: 'none' }
    if (!input.outerResized) return { action: 'none' }
  }
  return { action: 'refit', zoom: input.fitZoom }
}

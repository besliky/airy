/**
 * Split-view scroll-sync math (UX-1704). Pure and DOM-free so the proportion
 * and the loop protection are unit-testable without a live editor or frame.
 *
 * The sync is proportional, not structural: both panes map their own scroll
 * range to a 0..1 progress value and the other pane jumps to the same
 * progress. Good enough for "where am I in the document", cheap on every
 * scroll event, and independent of line counts (wrapping changes them).
 */

/** the scrollable state of one pane (a CodeMirror scrollDOM or the preview viewport) */
export interface ScrollerState {
  scrollTop: number
  scrollHeight: number
  clientHeight: number
}

/** normalized 0..1 progress of a scrollable viewport; 0 when it cannot scroll */
export function scrollRatio(s: ScrollerState): number {
  const max = s.scrollHeight - s.clientHeight
  if (max <= 0) return 0
  return Math.min(1, Math.max(0, s.scrollTop / max))
}

/** absolute scrollTop for a ratio, clamped into the viewport's legal range */
export function ratioToScrollTop(ratio: number, s: ScrollerState): number {
  const max = Math.max(0, s.scrollHeight - s.clientHeight)
  const r = Math.min(1, Math.max(0, ratio))
  return r * max
}

/** below this px difference a reposition is noise: applying it would itself be the start of a feedback hum */
export const SYNC_STEP_PX = 1

/** true when the two positions differ enough to warrant a programmatic scroll */
export function isSyncSignificant(from: number, to: number, epsilon = SYNC_STEP_PX): boolean {
  return Math.abs(to - from) >= epsilon
}

/** ratio deltas below this are scrollbar rounding, not a user scroll */
export const SYNC_RATIO_EPSILON = 0.002

/** how long a programmatically scrolled pane's own reports stay silenced */
export const SYNC_ECHO_WINDOW_MS = 150

/** which pane was (or is being) driven: the editor source or the preview frame */
export type SyncSide = 'source' | 'preview'

/**
 * Echo suppression for scroll-sync: driving pane B (because pane A scrolled)
 * silences B's own scroll reports for a short window, so the programmatic
 * reposition cannot bounce back as a fresh user scroll and the two panes
 * ping-pong forever. Timestamp-based, with wall-clock `now` injected so the
 * window logic is unit-testable without fake timers.
 */
export class ScrollSyncGate {
  private drivenAt: Record<SyncSide, number> = { source: -Infinity, preview: -Infinity }

  constructor(private readonly windowMs: number = SYNC_ECHO_WINDOW_MS) {}

  /** a programmatic scroll was sent to `side`: silence its echoes from now on */
  drive(side: SyncSide, now: number): void {
    this.drivenAt[side] = now
  }

  /** true while `side` may still be ringing from a programmatic reposition */
  shouldIgnore(side: SyncSide, now: number): boolean {
    return now - this.drivenAt[side] < this.windowMs
  }
}

/**
 * Pure logic for slide shows: playback sequence computation + rehearsal timing accumulation.
 * Extracted from SlideShowView for unit testing (no React/DOM dependency).
 */

/** Custom show: subset of slides in user-specified order (original indexes). App persists it per document to localStorage. */
export interface CustomShow {
  id: string
  name: string
  slideIndices: number[]
}

/**
 * Compute the playback sequence (array of original indexes).
 * - Default: all slides in order, skipping hidden ones (starting from a hidden slide still plays it)
 * - Non-empty customOrder: play in its order (out-of-range slides filtered; hidden slides still skipped, except the start slide)
 * - Fallback: when the result is empty, at least play the start slide
 */
export function computePlayOrder(
  slides: ReadonlyArray<{ hidden?: boolean }>,
  startAt: number,
  customOrder?: readonly number[],
): number[] {
  const playable = (i: number) => slides[i] != null && (!slides[i]!.hidden || i === startAt)
  const o =
    customOrder && customOrder.length > 0
      ? customOrder.filter(playable)
      : slides.map((_, i) => i).filter(playable)
  return o.length > 0 ? [...o] : [startAt]
}

// ── Rehearsal timing ─────────────────────────────────────────────────────────────

/** Rehearsal timing state: perPageMs accumulates dwell milliseconds by original slide index. */
export interface RehearseTiming {
  perPageMs: number[]
  /** Slide currently dwelt on (original index; -1 = finished) */
  currentIndex: number
  /** Timestamp of entering the current slide (ms) */
  enteredAt: number
}

/** Start rehearsal: begin timing from startIndex. */
export function startRehearse(slideCount: number, startIndex: number, now: number): RehearseTiming {
  return {
    perPageMs: new Array(Math.max(0, slideCount)).fill(0),
    currentIndex: startIndex,
    enteredAt: now,
  }
}

/** Page turn: accumulate the current slide's dwell into perPageMs, then switch to nextIndex and restart timing (revisiting a slide keeps accumulating). */
export function switchRehearsePage(
  t: RehearseTiming,
  nextIndex: number,
  now: number,
): RehearseTiming {
  const perPageMs = t.perPageMs.slice()
  if (t.currentIndex >= 0 && t.currentIndex < perPageMs.length) {
    perPageMs[t.currentIndex]! += Math.max(0, now - t.enteredAt)
  }
  return { perPageMs, currentIndex: nextIndex, enteredAt: now }
}

/**
 * End rehearsal: accumulate the last slide's dwell and keep it in milliseconds
 * (the exact dwell is what gets written to advTm — UX-1768; second rounding is
 * display-only, see formatClock). Unvisited slides stay 0.
 */
export function finishRehearse(t: RehearseTiming, now: number): number[] {
  const final = switchRehearsePage(t, -1, now)
  return final.perPageMs.map((ms) => (ms > 0 ? ms : 0))
}

/** m:ss clock display (rehearsal timer bar / save confirmation dialog). */
export function formatClock(ms: number): string {
  const sec = Math.max(0, Math.floor(ms / 1000))
  return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, '0')}`
}

// ── Record Slide Show session ────────────────────────────────────────────────────
// A Record Slide Show (PowerPoint parity minus narration — see PAR-314) is an
// explicit recording session on top of the rehearsal clock: dwell accumulates
// only while the session is recording, pause/resume freezes the clock without
// ending the session, and the accumulated dwell is saved as auto-advance
// timings (<p:transition advTm>) exactly like a rehearsal.

export type RecordPhase = 'recording' | 'paused'

/** Record session state: rehearsal clock + whether dwell currently accumulates. */
export interface RecordSession {
  perPageMs: number[]
  /** Slide currently dwelt on (original index; -1 = stopped) */
  currentIndex: number
  /** Timestamp the current (unaccumulated) dwell window started at (ms) */
  enteredAt: number
  phase: RecordPhase
}

/** Start a Record Slide Show session: recording immediately from startIndex. */
export function startRecord(slideCount: number, startIndex: number, now: number): RecordSession {
  return {
    perPageMs: new Array(Math.max(0, slideCount)).fill(0),
    currentIndex: startIndex,
    enteredAt: now,
    phase: 'recording',
  }
}

/** Pause recording: bank the current dwell; page turns no longer accumulate until resumed. */
export function pauseRecord(t: RecordSession, now: number): RecordSession {
  if (t.phase === 'paused') return t
  const banked = switchRecordPage(t, t.currentIndex, now)
  return { ...banked, phase: 'paused' }
}

/** Resume recording: a fresh dwell window starts for the current slide. */
export function resumeRecord(t: RecordSession, now: number): RecordSession {
  if (t.phase === 'recording') return t
  return { ...t, phase: 'recording', enteredAt: now }
}

/**
 * Page turn during a record session: accumulate the current slide's dwell while
 * recording (a paused session moves on without accumulating), then switch to
 * nextIndex (revisiting a slide keeps accumulating, as in a rehearsal).
 */
export function switchRecordPage(t: RecordSession, nextIndex: number, now: number): RecordSession {
  const perPageMs = t.perPageMs.slice()
  if (t.phase === 'recording' && t.currentIndex >= 0 && t.currentIndex < perPageMs.length) {
    perPageMs[t.currentIndex]! += Math.max(0, now - t.enteredAt)
  }
  return { perPageMs, currentIndex: nextIndex, enteredAt: now, phase: t.phase }
}

/**
 * Stop the session: bank the last dwell while recording (a paused session keeps
 * only what was banked at pause), then keep the per-slide dwell in milliseconds
 * exactly as accumulated — PowerPoint stores advTm with ms precision, so no
 * second rounding here (UX-1768; the HUD formats via formatClock instead).
 * Unvisited slides stay 0.
 */
export function finishRecord(t: RecordSession, now: number): number[] {
  const final = t.phase === 'recording' ? switchRecordPage(t, -1, now) : { ...t, currentIndex: -1 }
  return final.perPageMs.map((ms) => (ms > 0 ? ms : 0))
}

/** Dwell of the current slide so far (frozen while paused) — record HUD clock. */
export function currentRecordMs(t: RecordSession, now: number): number {
  const sinceEntered = t.phase === 'recording' ? Math.max(0, now - t.enteredAt) : 0
  return (t.perPageMs[t.currentIndex] ?? 0) + sinceEntered
}

/** Total recorded time so far (frozen while paused) — record HUD progress. */
export function totalRecordMs(t: RecordSession, now: number): number {
  const sinceEntered = t.phase === 'recording' ? Math.max(0, now - t.enteredAt) : 0
  return t.perPageMs.reduce((a, b) => a + b, 0) + sinceEntered
}

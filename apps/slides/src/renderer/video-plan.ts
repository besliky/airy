/**
 * Pure timeline planning for File > Export Video — turns the deck (hidden
 * flags), rehearsed auto-advance times and per-slide transitions into a
 * playback timeline the canvas recorder walks frame by frame. No DOM, so
 * every rule here (hidden skip, timing fallbacks, crossfade durations) is
 * unit-testable.
 *
 * Semantics (documented in the export dialog too):
 * - Hidden slides are skipped exactly like images/PDF export and the show.
 * - Slide dwell: the slide's rehearsed auto-advance time (advTm); a slide
 *   without one (or with timings disabled) falls back to the dialog's
 *   seconds-per-slide, like PowerPoint's Create a Video default.
 * - Transition: the INCOMING slide's transition (same ownership as the show
 *   view) plays as a crossfade after the previous slide's dwell — fade time
 *   is added on top of the dwells, not eaten out of them. 'none' is a hard
 *   cut; morph degrades to a crossfade (the show does the same when there is
 *   nothing to tween from); 'random' resolves once at plan time so the plan
 *   is stable.
 */
import type { TransitionKind } from '../shared/ipc'
import { resolveRandomTransition } from './transition-play'

/** Effect default durations (ms) — the CSS keyframes' defaults in styles.css. */
export const VIDEO_TRANSITION_DEFAULT_MS: Record<TransitionKind, number> = {
  none: 0,
  morph: 400,
  fade: 500,
  push: 500,
  wipe: 600,
  split: 550,
  circle: 600,
  cover: 500,
  pull: 500,
  dissolve: 600,
  zoom: 450,
  random: 500,
}

/** One visible slide's stretch of the video. */
export interface VideoTimelineItem {
  /** Original deck index of the slide (hidden slides never appear) */
  slideIndex: number
  /** How long the slide holds at full opacity (ms) before its outgoing fade */
  holdMs: number
  /** Crossfade into the NEXT slide (ms); 0 = hard cut / last slide */
  fadeOutMs: number
}

/** Timeline plan: consecutive items; totalMs = sum(holdMs + fadeOutMs). */
export interface VideoTimeline {
  items: VideoTimelineItem[]
  totalMs: number
}

export interface VideoPlanOptions {
  fps: number
  /** Use rehearsed auto-advance times; slides without one still fall back */
  useTimings: boolean
  /** Fallback dwell for slides without (or ignoring) timings, in seconds */
  secondsPerSlide: number
  /** Render transitions as crossfades; false = hard cuts everywhere */
  includeTransitions: boolean
}

/** Deck facts the planner needs (RenderSlide satisfies this structurally). */
export type VideoPlanSlide = { hidden?: boolean }

export interface VideoPlanInput {
  slides: ReadonlyArray<VideoPlanSlide>
  /** Rehearsed auto-advance ms per deck index (null/0 = none recorded) */
  advanceMs: ReadonlyArray<number | null>
  /** Per-slide transition spec (deck order) */
  transitions: ReadonlyArray<{ kind: TransitionKind; durationMs: number | null }>
  options: VideoPlanOptions
}

/** Crossfade duration entering the slide at deck index `nextIndex` (its own transition), 0 for a cut. */
function fadeOutMsOf(input: VideoPlanInput, nextIndex: number, rand: () => number): number {
  if (!input.options.includeTransitions) return 0
  const spec = input.transitions[nextIndex]
  if (!spec || spec.kind === 'none') return 0
  let kind: TransitionKind = spec.kind
  if (kind === 'random') kind = resolveRandomTransition(rand)
  return Math.max(0, spec.durationMs ?? VIDEO_TRANSITION_DEFAULT_MS[kind])
}

/** Dwell of visible slide at `pos`: rehearsed time, else the dialog fallback. */
function holdMsOf(input: VideoPlanInput, pos: number): number {
  const rehearsed = input.options.useTimings ? input.advanceMs[pos] : null
  return rehearsed && rehearsed > 0 ? rehearsed : input.options.secondsPerSlide * 1000
}

/**
 * Build the export timeline. Hidden slides are dropped; every visible slide
 * holds for its dwell and then crossfades into the next visible slide with
 * that slide's transition duration. An empty deck (everything hidden / no
 * slides) yields an empty timeline.
 */
export function buildVideoTimeline(
  input: VideoPlanInput,
  rand: () => number = Math.random,
): VideoTimeline {
  const visible = input.slides
    .map((slide, index) => ({ slide, index }))
    .filter(({ slide }) => !slide.hidden)
  const items: VideoTimelineItem[] = visible.map(({ index }, pos) => ({
    slideIndex: index,
    holdMs: holdMsOf(input, index),
    // the outgoing fade is the NEXT visible slide's own transition (deck-indexed)
    fadeOutMs: pos < visible.length - 1 ? fadeOutMsOf(input, visible[pos + 1]!.index, rand) : 0,
  }))
  return { items, totalMs: items.reduce((sum, it) => sum + it.holdMs + it.fadeOutMs, 0) }
}

/** What one output frame shows: `from` fully, `to` over it at alpha (0 = same slide). */
export interface VideoFrameState {
  /** Item index in the timeline being left (or shown) */
  from: number
  /** Item index coming in (=== from outside a fade) */
  to: number
  /** Crossfade progress 0..1 (0 = only `from`) */
  alpha: number
}

/**
 * Sample the timeline at `tMs` (clamped to [0, totalMs]). Inside a fade
 * window: from/to with the progress; inside a hold: from === to, alpha 0.
 */
export function sampleTimeline(timeline: VideoTimeline, tMs: number): VideoFrameState {
  const { items } = timeline
  if (items.length === 0) return { from: 0, to: 0, alpha: 0 }
  let t = Math.max(0, tMs)
  for (let i = 0; i < items.length; i++) {
    const it = items[i]!
    if (t < it.holdMs) return { from: i, to: i, alpha: 0 }
    t -= it.holdMs
    if (it.fadeOutMs > 0 && t < it.fadeOutMs) {
      return { from: i, to: i + 1, alpha: Math.min(1, t / it.fadeOutMs) }
    }
    t -= it.fadeOutMs
  }
  // past the end: the last slide holds forever (the recorder's trailing frames)
  return { from: items.length - 1, to: items.length - 1, alpha: 0 }
}

/** Total output frame count at the given fps (>= 1 when the deck has content). */
export function videoFrameCount(timeline: VideoTimeline, fps: number): number {
  if (timeline.items.length === 0) return 0
  return Math.max(1, Math.round((timeline.totalMs * fps) / 1000))
}

/**
 * Output pixel size for a resolution preset: the preset names the SLIDE
 * height (720/1080); width follows the slide aspect (16:9 decks get the
 * familiar 1280x720 / 1920x1080, 4:3 decks 960x720 / 1440x1080). Both axes
 * round to even — H.264/VP9 require macroblock-aligned dimensions.
 */
export function videoFrameDimensions(
  widthPx: number,
  heightPx: number,
  heightPreset: 720 | 1080,
): { width: number; height: number } {
  const even = (n: number) => Math.max(2, Math.round(n / 2) * 2)
  const height = even(heightPreset)
  return { width: even((widthPx / heightPx) * height), height }
}

/**
 * Aspect-fit a slide into the output frame: uniform min-scale, centered — a
 * deck with mixed slide sizes (a 4:3 slide among 16:9) letterbox/pillarboxes
 * each slide into the frame instead of stretching it to the first slide's
 * shape (BUG-1211). A matching aspect fills the whole frame.
 */
export function fitIntoFrame(
  slideW: number,
  slideH: number,
  frameW: number,
  frameH: number,
): { dx: number; dy: number; dw: number; dh: number } {
  const scale = Math.min(frameW / slideW, frameH / slideH)
  const dw = slideW * scale
  const dh = slideH * scale
  return { dx: Math.floor((frameW - dw) / 2), dy: Math.floor((frameH - dh) / 2), dw, dh }
}

/** True when at least one slide carries a rehearsed auto-advance time. */
export function hasRehearseTimings(advanceMs: ReadonlyArray<number | null>): boolean {
  return advanceMs.some((ms) => ms != null && ms > 0)
}

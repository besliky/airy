/**
 * Pure transition playback mapping — shared by the show view, the audience window
 * and the editor's one-shot canvas preview.
 *
 * A slide's TransitionSpec (read from <p:transition>) resolves to a CSS animation
 * plan: the frame class `ss-anim-<kind>` (or `tp-<kind>` in the editor preview —
 * same keyframes family), a shared direction modifier class `tr-<dir>` that feeds
 * the keyframes' CSS custom properties, and an optional explicit duration
 * (p14:dur) overriding the CSS default. 'random' resolves to one of the animated
 * kinds; 'morph' either tweens via MorphStage or degrades to fade.
 */
import type { TransitionKind, TransitionSpec } from '../shared/ipc'
import { TRANSITION_DIRS } from '../shared/ipc'
/** Kinds with a CSS approximation (everything but none/random/morph). */
export const ANIMATED_TRANSITIONS: readonly TransitionKind[] = [
  'fade',
  'push',
  'wipe',
  'split',
  'circle',
  'cover',
  'pull',
  'dissolve',
  'zoom',
] as const

/** Resolve 'random' to a concrete animated kind (stable per call). */
export function resolveRandomTransition(rand: () => number): TransitionKind {
  return ANIMATED_TRANSITIONS[Math.floor(rand() * ANIMATED_TRANSITIONS.length)]!
}

/** Direction modifier class for the animated element ('' = the kind's default direction). */
export function transitionDirClass(kind: TransitionKind, spec: TransitionSpec): string {
  if (spec.dir == null || !TRANSITION_DIRS[kind].includes(spec.dir)) return ''
  if (kind === 'split') {
    // split's dir is in/out; only the vertical variants change the visual axis
    return spec.orient === 'vert' ? ' tr-split-vert' : ''
  }
  if (kind === 'zoom') {
    return spec.dir === 'out' ? ' tr-zoom-out' : ''
  }
  return ` tr-${spec.dir}`
}

/** One planned page-turn animation. */
export interface TransitionPlayback {
  /** Concrete kind (random/morph already resolved) */
  kind: TransitionKind
  /** CSS classes for the animated frame ('' = no animation) */
  css: string
  /** Explicit duration in ms (inline animationDuration); null = CSS default */
  durationMs: number | null
  /** Morph tween requested (MorphStage handles it; css stays empty) */
  morph: boolean
}

/**
 * Resolve a slide's transition into a playback plan. canMorph=false (first page,
 * same page) degrades morph to a fade, matching the PowerPoint behavior of
 * falling back when there is nothing to tween from.
 */
export function planTransition(
  spec: TransitionSpec,
  opts: { canMorph: boolean; random: () => number },
): TransitionPlayback {
  let kind: TransitionKind = spec.kind
  if (kind === 'random') kind = resolveRandomTransition(opts.random)
  let morph = false
  if (kind === 'morph') {
    if (opts.canMorph) morph = true
    else kind = 'fade'
  }
  const css = kind !== 'none' && !morph ? `ss-anim-${kind}${transitionDirClass(kind, spec)}` : ''
  return { kind, css, durationMs: spec.durationMs, morph }
}

/**
 * Switch a spec's kind (gallery click), keeping the Effect Options that still
 * apply: the direction survives when the new kind offers it, the duration always
 * carries over (PowerPoint keeps duration when switching effects).
 */
export function transitionWithKind(spec: TransitionSpec, kind: TransitionKind): TransitionSpec {
  const dir = spec.dir != null && TRANSITION_DIRS[kind].includes(spec.dir) ? spec.dir : undefined
  return {
    kind,
    ...(dir != null ? { dir } : {}),
    ...(kind === 'split' ? { orient: spec.orient ?? 'horz' } : {}),
    durationMs: spec.durationMs,
  }
}

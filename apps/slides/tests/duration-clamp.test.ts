import { describe, expect, it } from 'vitest'
import { clampDurationSeconds } from '../src/renderer/components/ribbon-shared'

/**
 * UX-1105: the transition / animation duration inputs declare 0.1–60 / 0–60
 * in their min/max attributes, but the blur commit used to accept any
 * positive number — a typed «999» became 999 000 ms with no warning. The
 * blur handlers now clamp through this helper before rounding to
 * milliseconds, so the committed value always lands in the declared range
 * (the input re-mounts via its durationMs key and shows the clamped value).
 */
describe('clampDurationSeconds (UX-1105)', () => {
  it.each([
    ['below the floor is lifted to it', 0.05, 0.1, 60, 0.1],
    ['the floor itself passes', 0.1, 0.1, 60, 0.1],
    ['a regular value passes through', 2.35, 0.1, 60, 2.35],
    ['the ceiling itself passes', 60, 0.1, 60, 60],
    ['above the ceiling is cut to it', 999, 0.1, 60, 60],
    ['the animation floor is zero, negatives clamp up', -5, 0, 60, 0],
  ] as const)('%s', (_desc, v, min, max, want) => {
    expect(clampDurationSeconds(v, min, max)).toBe(want)
  })
})

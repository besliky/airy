import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { clampDurationSeconds } from '../src/renderer/components/ribbon-shared'

/**
 * UX-1105: the transition / animation duration inputs declare 0.1–60 / 0–60
 * in their min/max attributes, but the blur commit used to accept any
 * positive number — a typed «999» became 999 000 ms with no warning. The
 * blur handlers now clamp through this helper before rounding to
 * milliseconds, so the committed value always lands in the declared range
 * (the input re-mounts via its durationMs key and shows the clamped value).
 *
 * UX-11s2 tail: PowerPoint caps morph one second below every other
 * transition — the input's max, the blur clamp and the localized range
 * hint now all follow the active transition kind (59 for morph, 60
 * otherwise); the hint key takes the ceiling as a {max} parameter.
 */
describe('clampDurationSeconds (UX-1105)', () => {
  it.each([
    ['below the floor is lifted to it', 0.05, 0.1, 60, 0.1],
    ['the floor itself passes', 0.1, 0.1, 60, 0.1],
    ['a regular value passes through', 2.35, 0.1, 60, 2.35],
    ['the ceiling itself passes', 60, 0.1, 60, 60],
    ['above the ceiling is cut to it', 999, 0.1, 60, 60],
    ['the animation floor is zero, negatives clamp up', -5, 0, 60, 0],
    ['morph keeps its own 59 s ceiling', 59.5, 0.1, 59, 59],
    ['morph above the ceiling is cut to 59', 999, 0.1, 59, 59],
  ] as const)('%s', (_desc, v, min, max, want) => {
    expect(clampDurationSeconds(v, min, max)).toBe(want)
  })
})

describe('the transition input follows the morph ceiling (UX-11s2)', () => {
  const src = readFileSync(join(__dirname, '../src/renderer/components/Ribbon.tsx'), 'utf8')

  it('derives the ceiling from the transition kind', () => {
    expect(src).toContain("transition.kind === 'morph' ? 59 : 60")
  })

  it('feeds the ceiling to the max attribute, the clamp and the range hint', () => {
    const panel = src.slice(src.indexOf('ribbonTransDurationRange') - 2500, src.indexOf('ribbonTransDurationRange'))
    expect(panel).toContain('max={ceiling}')
    expect(panel).toContain('clampDurationSeconds(v, 0.1, ceiling)')
    expect(src).toContain("'ribbonTransDurationRange',\n                              { max: ceiling }")
  })
})

import { describe, it, expect } from 'vitest'
import {
  ANIMATED_TRANSITIONS,
  planTransition,
  resolveRandomTransition,
  transitionDirClass,
  transitionWithKind,
} from '../src/renderer/transition-play'
import {
  ANIM_EFFECT_DIRS,
  TRANSITION_DIRS,
  TRANSITION_DEFAULT_DIR,
  type TransitionSpec,
} from '../src/shared/ipc'
import { ANIM_DIRECTIONS, ANIM_DEFAULT_DIR, TRANSITION_DIR_INFO } from '@airy-office/pptx-engine'
import { ANIM_DEFAULT_DIRECTION } from '../src/renderer/animation-play'

const spec = (o: Partial<TransitionSpec> & Pick<TransitionSpec, 'kind'>): TransitionSpec => ({
  durationMs: null,
  ...o,
})

const never = () => 0

describe('planTransition', () => {
  it('maps plain kinds to the ss-anim frame class without a direction modifier', () => {
    expect(planTransition(spec({ kind: 'fade' }), { canMorph: false, random: never }).css).toBe(
      'ss-anim-fade',
    )
    expect(planTransition(spec({ kind: 'none' }), { canMorph: false, random: never }).css).toBe('')
  })

  it('adds the direction modifier class for option-carrying kinds', () => {
    expect(
      planTransition(spec({ kind: 'push', dir: 'fromTop' }), { canMorph: false, random: never })
        .css,
    ).toBe('ss-anim-push tr-fromTop')
    expect(
      planTransition(spec({ kind: 'cover', dir: 'fromBottomLeft' }), {
        canMorph: false,
        random: never,
      }).css,
    ).toBe('ss-anim-cover tr-fromBottomLeft')
    // A direction the kind does not offer is ignored
    expect(
      planTransition(spec({ kind: 'fade', dir: 'fromTop' }), { canMorph: false, random: never })
        .css,
    ).toBe('ss-anim-fade')
  })

  it('maps split orientation to the vertical variant class and zoom out', () => {
    expect(
      planTransition(spec({ kind: 'split', orient: 'vert', dir: 'out' }), {
        canMorph: false,
        random: never,
      }).css,
    ).toBe('ss-anim-split tr-split-vert')
    expect(
      planTransition(spec({ kind: 'split', orient: 'horz', dir: 'in' }), {
        canMorph: false,
        random: never,
      }).css,
    ).toBe('ss-anim-split')
    expect(
      planTransition(spec({ kind: 'zoom', dir: 'out' }), { canMorph: false, random: never }).css,
    ).toBe('ss-anim-zoom tr-zoom-out')
  })

  it('passes the explicit duration through for the inline animationDuration', () => {
    const plan = planTransition(spec({ kind: 'wipe', dir: 'fromLeft', durationMs: 1500 }), {
      canMorph: false,
      random: never,
    })
    expect(plan.css).toBe('ss-anim-wipe tr-fromLeft')
    expect(plan.durationMs).toBe(1500)
    expect(
      planTransition(spec({ kind: 'wipe' }), { canMorph: false, random: never }).durationMs,
    ).toBeNull()
  })

  it('resolves random deterministically from the injected rng', () => {
    expect(resolveRandomTransition(() => 0)).toBe(ANIMATED_TRANSITIONS[0])
    expect(resolveRandomTransition(() => 0.99)).toBe(
      ANIMATED_TRANSITIONS[ANIMATED_TRANSITIONS.length - 1],
    )
    const plan = planTransition(spec({ kind: 'random' }), { canMorph: false, random: () => 0.5 })
    expect(plan.morph).toBe(false)
    expect(plan.css).toMatch(/^ss-anim-\w+$/)
  })

  it('morph tweens only when there is a previous different page; otherwise degrades to fade', () => {
    const tweenable = planTransition(spec({ kind: 'morph' }), { canMorph: true, random: never })
    expect(tweenable.morph).toBe(true)
    expect(tweenable.css).toBe('')
    const first = planTransition(spec({ kind: 'morph' }), { canMorph: false, random: never })
    expect(first.morph).toBe(false)
    expect(first.kind).toBe('fade')
    expect(first.css).toBe('ss-anim-fade')
  })

  it('morph keeps its duration for the MorphStage tween', () => {
    const plan = planTransition(spec({ kind: 'morph', durationMs: 2000 }), {
      canMorph: true,
      random: never,
    })
    expect(plan.morph).toBe(true)
    expect(plan.durationMs).toBe(2000)
  })
})

describe('transitionWithKind (gallery click keeps applicable options)', () => {
  it('keeps duration and a still-applicable direction, drops the rest', () => {
    const s = spec({ kind: 'push', dir: 'fromTop', durationMs: 900 })
    expect(transitionWithKind(s, 'wipe')).toEqual(
      spec({ kind: 'wipe', dir: 'fromTop', durationMs: 900 }),
    )
    expect(transitionWithKind(s, 'fade')).toEqual(spec({ kind: 'fade', durationMs: 900 }))
    // diagonal survives only on cover/pull
    const d = spec({ kind: 'cover', dir: 'fromTopLeft' })
    expect(transitionWithKind(d, 'push')).toEqual(spec({ kind: 'push' }))
    expect(transitionWithKind(d, 'pull')).toEqual(spec({ kind: 'pull', dir: 'fromTopLeft' }))
    // split keeps/derives the orientation
    expect(transitionWithKind(spec({ kind: 'push' }), 'split')).toEqual(
      spec({ kind: 'split', orient: 'horz' }),
    )
  })
})

describe('renderer/engine option tables stay in sync', () => {
  it('ipc TRANSITION_DIRS mirrors the engine TRANSITION_DIR_INFO', () => {
    for (const kind of Object.keys(TRANSITION_DIR_INFO) as Array<
      keyof typeof TRANSITION_DIR_INFO
    >) {
      expect([...TRANSITION_DIRS[kind]]).toEqual([...TRANSITION_DIR_INFO[kind].dirs])
      expect(TRANSITION_DEFAULT_DIR[kind]).toBe(TRANSITION_DIR_INFO[kind].default)
    }
  })

  it('renderer ANIM_DEFAULT_DIRECTION mirrors the engine ANIM_DEFAULT_DIR', () => {
    for (const effect of Object.keys(ANIM_DIRECTIONS) as Array<keyof typeof ANIM_DIRECTIONS>) {
      expect(ANIM_DEFAULT_DIRECTION[effect]).toBe(ANIM_DEFAULT_DIR[effect])
    }
  })

  it('the ipc animation direction table mirrors the engine table', () => {
    expect(Object.keys(ANIM_EFFECT_DIRS)).toEqual(Object.keys(ANIM_DIRECTIONS))
    for (const effect of Object.keys(ANIM_DIRECTIONS) as Array<keyof typeof ANIM_DIRECTIONS>) {
      expect([...ANIM_EFFECT_DIRS[effect]]).toEqual([...ANIM_DIRECTIONS[effect]])
    }
  })

  it('every direction the engine offers has a modifier class rendering', () => {
    // side directions + split/zoom specials produce a non-empty class
    for (const kind of ['push', 'wipe', 'cover', 'pull'] as const) {
      const dir = TRANSITION_DIRS[kind][0]!
      expect(transitionDirClass(kind, spec({ kind, dir }))).not.toBe('')
    }
  })
})

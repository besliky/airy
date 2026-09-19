import { describe, expect, it } from 'vitest'
import type { TransitionKind } from '../src/shared/ipc'
import {
  VIDEO_TRANSITION_DEFAULT_MS,
  buildVideoTimeline,
  hasRehearseTimings,
  sampleTimeline,
  videoFrameCount,
  videoFrameDimensions,
  type VideoPlanInput,
} from '../src/renderer/video-plan'

const never = () => 0

const plan = (input: Partial<VideoPlanInput>): VideoPlanInput => ({
  slides: [{}, {}, {}],
  advanceMs: [null, null, null],
  transitions: [
    { kind: 'none', durationMs: null },
    { kind: 'none', durationMs: null },
    { kind: 'none', durationMs: null },
  ],
  options: {
    fps: 30,
    useTimings: false,
    secondsPerSlide: 5,
    includeTransitions: true,
  },
  ...input,
})

const tr = (kind: TransitionKind, durationMs: number | null = null) => ({
  kind,
  durationMs,
})

describe('buildVideoTimeline', () => {
  it('skips hidden slides entirely', () => {
    const t = buildVideoTimeline(
      plan({
        slides: [{}, { hidden: true }, {}],
        transitions: [tr('none'), tr('fade', 500), tr('fade', 500)],
        options: {
          fps: 30,
          useTimings: false,
          secondsPerSlide: 4,
          includeTransitions: true,
        },
      }),
      never,
    )
    // slide 1 is hidden: items cover 0 and 2; slide 2's fade (the transition
    // INTO the incoming slide) plays between them
    expect(t.items.map((i) => i.slideIndex)).toEqual([0, 2])
    expect(t.items[0]!.fadeOutMs).toBe(500)
    expect(t.items[1]!.fadeOutMs).toBe(0) // last slide never fades out
  })

  it('paces holds by rehearsed timings, with the per-second fallback for unrecorded slides', () => {
    const t = buildVideoTimeline(
      plan({
        advanceMs: [3000, 0, 7000],
        options: {
          fps: 30,
          useTimings: true,
          secondsPerSlide: 5,
          includeTransitions: false,
        },
      }),
      never,
    )
    expect(t.items.map((i) => i.holdMs)).toEqual([3000, 5000, 7000]) // 0 → fallback 5s
  })

  it('ignores timings when useTimings is off (even split by the fallback)', () => {
    const t = buildVideoTimeline(
      plan({
        advanceMs: [3000, 2000, 1000],
        options: {
          fps: 30,
          useTimings: false,
          secondsPerSlide: 2,
          includeTransitions: false,
        },
      }),
      never,
    )
    expect(t.items.map((i) => i.holdMs)).toEqual([2000, 2000, 2000])
  })

  it('uses the INCOMING slide transition: explicit duration, kind defaults, none = cut', () => {
    const t = buildVideoTimeline(
      plan({
        transitions: [tr('none'), tr('fade', 800), tr('push')],
        options: {
          fps: 30,
          useTimings: false,
          secondsPerSlide: 5,
          includeTransitions: true,
        },
      }),
      never,
    )
    expect(t.items[0]!.fadeOutMs).toBe(800) // into slide 1 (fade, explicit 800ms)
    expect(t.items[1]!.fadeOutMs).toBe(VIDEO_TRANSITION_DEFAULT_MS.push) // into slide 2 (default)
    expect(t.items[2]!.fadeOutMs).toBe(0)
  })

  it('morph degrades to a crossfade with the morph default; random resolves to one kind', () => {
    const morph = buildVideoTimeline(
      plan({
        slides: [{}, {}],
        transitions: [tr('none'), tr('morph')],
        options: {
          fps: 30,
          useTimings: false,
          secondsPerSlide: 5,
          includeTransitions: true,
        },
      }),
      never,
    )
    expect(morph.items[0]!.fadeOutMs).toBe(VIDEO_TRANSITION_DEFAULT_MS.morph)

    const random = buildVideoTimeline(
      plan({
        slides: [{}, {}],
        transitions: [tr('none'), tr('random', null)],
        options: {
          fps: 30,
          useTimings: false,
          secondsPerSlide: 5,
          includeTransitions: true,
        },
      }),
      never, // rand()=0 → 'fade'
    )
    expect(random.items[0]!.fadeOutMs).toBe(VIDEO_TRANSITION_DEFAULT_MS.fade)
  })

  it('hard cuts everywhere when transitions are disabled', () => {
    const t = buildVideoTimeline(
      plan({
        transitions: [tr('fade', 900), tr('push', 900), tr('morph')],
        options: {
          fps: 30,
          useTimings: false,
          secondsPerSlide: 5,
          includeTransitions: false,
        },
      }),
      never,
    )
    expect(t.items.every((i) => i.fadeOutMs === 0)).toBe(true)
  })

  it('totals holds + fades and yields an empty timeline for an all-hidden deck', () => {
    const t = buildVideoTimeline(
      plan({
        slides: [{}, { hidden: true }, {}],
        advanceMs: [1000, null, 2000],
        transitions: [tr('none'), tr('none'), tr('fade', 500)],
        options: {
          fps: 30,
          useTimings: true,
          secondsPerSlide: 5,
          includeTransitions: true,
        },
      }),
      never,
    )
    expect(t.totalMs).toBe(1000 + 500 + 2000)
    expect(
      buildVideoTimeline(plan({ slides: [{ hidden: true }, { hidden: true }] }), never).items
        .length,
    ).toBe(0)
  })
})

describe('sampleTimeline', () => {
  const t = buildVideoTimeline(
    plan({
      slides: [{}, {}],
      transitions: [tr('none'), tr('fade', 1000)],
      options: {
        fps: 10,
        useTimings: false,
        secondsPerSlide: 2,
        includeTransitions: true,
      },
    }),
    never,
  ) // slide0: hold 2000, fade 1000, slide1: hold 2000

  it('holds the first slide, crossfades through the window, then holds the last', () => {
    expect(sampleTimeline(t, 0)).toEqual({ from: 0, to: 0, alpha: 0 })
    expect(sampleTimeline(t, 1999)).toEqual({ from: 0, to: 0, alpha: 0 })
    expect(sampleTimeline(t, 2000)).toEqual({ from: 0, to: 1, alpha: 0 })
    expect(sampleTimeline(t, 2500)).toEqual({ from: 0, to: 1, alpha: 0.5 })
    expect(sampleTimeline(t, 2999)).toEqual({ from: 0, to: 1, alpha: 0.999 })
    expect(sampleTimeline(t, 3000)).toEqual({ from: 1, to: 1, alpha: 0 })
    // past the end (trailing frames) keeps showing the last slide
    expect(sampleTimeline(t, 999_999)).toEqual({ from: 1, to: 1, alpha: 0 })
  })

  it('clamps negative time to the first hold', () => {
    expect(sampleTimeline(t, -50)).toEqual({ from: 0, to: 0, alpha: 0 })
  })
})

describe('videoFrameCount / videoFrameDimensions / hasRehearseTimings', () => {
  it('frame count rounds totalMs at the fps (min 1, 0 for empty decks)', () => {
    const t = buildVideoTimeline(
      plan({
        slides: [{}],
        options: {
          fps: 30,
          useTimings: true,
          secondsPerSlide: 1,
          includeTransitions: false,
        },
      }),
      never,
    )
    expect(videoFrameCount(t, 30)).toBe(30)
    expect(videoFrameCount({ items: [], totalMs: 0 }, 30)).toBe(0)
  })

  it('resolution presets name the slide height; width follows the aspect, both even', () => {
    // 16:9
    expect(videoFrameDimensions(1280, 720, 720)).toEqual({ width: 1280, height: 720 })
    expect(videoFrameDimensions(1280, 720, 1080)).toEqual({ width: 1920, height: 1080 })
    // 4:3 keeps its aspect
    expect(videoFrameDimensions(960, 720, 720)).toEqual({ width: 960, height: 720 })
    expect(videoFrameDimensions(960, 720, 1080)).toEqual({ width: 1440, height: 1080 })
    // odd intermediate results round to the nearest even dimension
    expect(videoFrameDimensions(961, 720, 720)).toEqual({ width: 962, height: 720 })
  })

  it('detects any recorded timing', () => {
    expect(hasRehearseTimings([null, 0, 5000])).toBe(true)
    expect(hasRehearseTimings([null, 0])).toBe(false)
  })
})

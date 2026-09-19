import { describe, expect, it } from 'vitest'
import {
  pickRecorderMime,
  recordVideoTimeline,
  videoBitrate,
  type RecorderHost,
} from '../src/renderer/video-export'
import { buildVideoTimeline } from '../src/renderer/video-plan'

/**
 * Mock-pipeline tests for the video export: the canvas/MediaRecorder host is
 * faked (jsdom has no real canvas capture), recording every draw call so the
 * tests pin the frame stream the encoder would see — which slides, in which
 * order, how the crossfade composes, and that cancel/stop stay clean.
 */

/** One drawable stand-in per slide; the ctx logs which image index was painted. */
const image = (tag: string) => ({ tag }) as unknown as CanvasImageSource

interface FakeHostState {
  draws: Array<{ tag: string; alpha: number }>
  requestedFrames: number
  chunks: Blob[]
  recorderCalls: string[]
}

function fakeHost(state: FakeHostState): RecorderHost {
  return {
    createCanvas(width, height) {
      const ctx = {
        canvas: { width, height },
        fillStyle: '#000',
        globalAlpha: 1,
        fillRect() {},
        drawImage(img: { tag: string }, _x?: number, _y?: number, _w?: number, _h?: number) {
          state.draws.push({ tag: img.tag, alpha: this.globalAlpha })
        },
      }
      return {
        ctx: ctx as unknown as CanvasRenderingContext2D,
        captureStream() {
          return {
            getVideoTracks: () => [
              {
                requestFrame: () => {
                  state.requestedFrames++
                },
              },
            ],
          }
        },
      }
    },
    createRecorder(_stream, recOpts) {
      return {
        start() {
          state.recorderCalls.push(`start:${recOpts.mimeType}`)
        },
        stop() {
          state.recorderCalls.push('stop')
          state.chunks.push(new Blob([new Uint8Array([1])], { type: recOpts.mimeType }))
        },
        ondataavailable() {},
        onstop(handler) {
          handler()
        },
      }
    },
    now: () => 0,
    // resolve timers immediately: the wall-clock pacing collapses, the frame
    // SEQUENCE is what these tests pin
    setTimeout(fn) {
      queueMicrotask(fn)
      return 0
    },
  }
}

const never = () => 0

/** 2 slides x 1s hold each + a 1s crossfade between them (10fps → 30 frames). */
const twoSlideTimeline = () =>
  buildVideoTimeline(
    {
      slides: [{}, {}],
      advanceMs: [1000, 1000],
      transitions: [
        { kind: 'none', durationMs: null },
        { kind: 'fade', durationMs: 1000 },
      ],
      options: {
        fps: 10,
        useTimings: true,
        secondsPerSlide: 5,
        includeTransitions: true,
      },
    },
    never,
  )

describe('pickRecorderMime', () => {
  it('prefers MP4 (H.264) when the build can mux it', () => {
    expect(pickRecorderMime(() => true)).toEqual({
      mimeType: 'video/mp4;codecs=avc1.640028',
      container: 'mp4',
    })
  })

  it('falls back to WebM when only WebM codecs are supported', () => {
    expect(pickRecorderMime((m) => m.startsWith('video/webm'))).toEqual({
      mimeType: 'video/webm;codecs=vp9',
      container: 'webm',
    })
  })

  it('walks the WebM ladder and returns null when nothing is supported', () => {
    expect(pickRecorderMime((m) => m === 'video/webm;codecs=vp8')).toEqual({
      mimeType: 'video/webm;codecs=vp8',
      container: 'webm',
    })
    expect(pickRecorderMime(() => false)).toBeNull()
  })

  it('scales the bitrate with the output height', () => {
    expect(videoBitrate(720)).toBe(6_000_000)
    expect(videoBitrate(1080)).toBe(10_000_000)
  })
})

describe('recordVideoTimeline', () => {
  it('pushes one frame per plan frame, ordered by slide, crossfading between holds', async () => {
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const progress: number[] = []
    const blob = await recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideImages: [image('s0'), image('s1')],
        mimeType: 'video/mp4;codecs=avc1.640028',
        onProgress: (done, total) => progress.push(done / total),
      },
      fakeHost(state),
    )
    // timeline: hold s0 1000ms (frames 0..9), fade 1000ms (10..19), hold s1 1000ms (20..29)
    expect(state.requestedFrames).toBe(30)
    expect(state.recorderCalls).toEqual(['start:video/mp4;codecs=avc1.640028', 'stop'])
    expect(blob).not.toBeNull()
    expect(blob!.type).toBe('video/mp4;codecs=avc1.640028')
    expect(progress[0]).toBeCloseTo(1 / 30)
    expect(progress.at(-1)).toBe(1)

    // first frame: only slide 0 at full opacity
    expect(state.draws[0]).toEqual({ tag: 's0', alpha: 1 })
    // last frame (frame 29, t=2900ms): only slide 1
    expect(state.draws.at(-1)).toEqual({ tag: 's1', alpha: 1 })
    // fade frames 11..19 draw the pair (frame 10's alpha is still 0 → s0 alone):
    // frame 15 (t=1500ms, alpha 0.5) is the pair at draws 19..20
    expect(state.draws[19]).toEqual({ tag: 's0', alpha: 1 })
    expect(state.draws[20]!.tag).toBe('s1')
    expect(state.draws[20]!.alpha).toBeCloseTo(0.5)
  })

  it('resolves null for an empty timeline and honors cooperative cancel', async () => {
    const empty = await recordVideoTimeline(
      {
        timeline: { items: [], totalMs: 0 },
        fps: 10,
        width: 320,
        height: 180,
        slideImages: [],
        mimeType: 'video/webm',
      },
      fakeHost({ draws: [], requestedFrames: 0, chunks: [], recorderCalls: [] }),
    )
    expect(empty).toBeNull()

    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const cancel = { current: true } // cancel before the first frame lands
    const cancelled = await recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideImages: [image('s0'), image('s1')],
        mimeType: 'video/webm',
        cancel,
      },
      fakeHost(state),
    )
    expect(cancelled).toBeNull()
    expect(state.requestedFrames).toBe(0)
    expect(state.recorderCalls).toEqual(['start:video/webm', 'stop'])
  })
})

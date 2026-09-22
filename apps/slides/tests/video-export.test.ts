import { describe, expect, it } from 'vitest'
import {
  pickRecorderMime,
  recordVideoTimeline,
  videoBitrate,
  type RecorderHost,
  type SlideBitmap,
  type SlideBitmapProvider,
  type VideoChunkSink,
} from '../src/renderer/video-export'
import { buildVideoTimeline } from '../src/renderer/video-plan'

/**
 * Mock-pipeline tests for the video export: the canvas/MediaRecorder host is
 * faked (jsdom has no real canvas capture), recording every draw call so the
 * tests pin the frame stream the encoder would see — which slides, in which
 * order, how the crossfade composes, and that cancel/stop stay clean. The
 * slide bitmaps come from a counting fake so the streaming window (BUG-1300)
 * is pinned here too: what got decoded, when it was released, and how the
 * recorder's chunks reached the sink.
 */

/** Bitmap stand-in per slide: a drawable tag plus a spied release. */
function slideBitmap(tag: string, log: { released: string[] }): SlideBitmap {
  return {
    image: { tag } as unknown as CanvasImageSource,
    release: () => {
      log.released.push(tag)
    },
  }
}

interface FakeHostState {
  draws: Array<{ tag: string; alpha: number }>
  requestedFrames: number
  chunks: Blob[]
  recorderCalls: string[]
  /** drawImage target boxes (x, y, w, h) — filled when a test passes slideSizes */
  drawBoxes?: Array<[number, number, number, number]>
}

/** Failure injections for the recorder host (BUG-1208 paths). */
interface FakeHostBehavior {
  /** fire onerror with this cause right after start (runtime encoder crash) */
  errorAfterStart?: Error
  /** swallow stop(): never fire onstop (encoder hangs the export) */
  hangOnStop?: boolean
  /** fire onstop but produce no chunks (empty recording) */
  noData?: boolean
}

function fakeHost(state: FakeHostState, behavior?: FakeHostBehavior): RecorderHost {
  return {
    createCanvas(width, height) {
      const ctx = {
        canvas: { width, height },
        fillStyle: '#000',
        globalAlpha: 1,
        fillRect() {},
        drawImage(img: { tag: string }, x?: number, y?: number, w?: number, h?: number) {
          state.draws.push({ tag: img.tag, alpha: this.globalAlpha })
          state.drawBoxes?.push([x ?? 0, y ?? 0, w ?? 0, h ?? 0])
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
      let errorHandler: ((e: { error?: unknown }) => void) | null = null
      let stopHandler: (() => void) | null = null
      let dataHandler: ((e: { data: Blob }) => void) | null = null
      return {
        start(timesliceMs?: number) {
          state.recorderCalls.push(`start:${recOpts.mimeType}:${timesliceMs}`)
          if (behavior?.errorAfterStart)
            queueMicrotask(() => errorHandler?.({ error: behavior.errorAfterStart }))
        },
        stop() {
          state.recorderCalls.push('stop')
          if (behavior?.hangOnStop) return
          if (!behavior?.noData) {
            // a real MediaRecorder flushes its buffer as a dataavailable just
            // before onstop — the pipeline collects chunks through that handler
            const chunk = new Blob([new Uint8Array([1])], { type: recOpts.mimeType })
            state.chunks.push(chunk)
            queueMicrotask(() => dataHandler?.({ data: chunk }))
          }
          queueMicrotask(() => stopHandler?.())
        },
        ondataavailable(handler) {
          dataHandler = handler
        },
        onstop(handler) {
          stopHandler = handler
        },
        onerror(handler) {
          errorHandler = handler
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

/** Sink fake: records the ordered chunk hand-off + how the stream settled. */
interface SinkTrace {
  writes: number[]
  finish: boolean[] // per FIRST settle: the aborted flag (finish is idempotent by contract)
}

function fakeSink(trace: SinkTrace): VideoChunkSink {
  let seq = 0
  let settled = false
  return {
    async write() {
      trace.writes.push(seq++)
    },
    async finish(aborted) {
      if (settled) return null // the real file sink settles once; so does the fake
      settled = true
      trace.finish.push(aborted)
      return aborted ? null : '/tmp/out.mp4'
    },
  }
}

/** Bitmap provider fake over per-slide tags; `log.released` fills via release(). */
function bitmapProvider(
  tags: string[],
  log: { released: string[]; fetched: string[] },
): SlideBitmapProvider {
  return async (itemIndex) => {
    const tag = tags[itemIndex]!
    log.fetched.push(tag)
    return slideBitmap(tag, log)
  }
}

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
    const log = { released: [] as string[], fetched: [] as string[] }
    const sinkTrace: SinkTrace = { writes: [], finish: [] }
    const recorded = await recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], log),
        sink: fakeSink(sinkTrace),
        mimeType: 'video/mp4;codecs=avc1.640028',
        onProgress: (done, total) => progress.push(done / total),
      },
      fakeHost(state),
    )
    // timeline: hold s0 1000ms (frames 0..9), fade 1000ms (10..19), hold s1 1000ms (20..29)
    expect(state.requestedFrames).toBe(30)
    expect(state.recorderCalls[0]).toBe('start:video/mp4;codecs=avc1.640028:1000')
    expect(state.recorderCalls[1]).toBe('stop')
    expect(recorded).toEqual({ bytes: 1, path: '/tmp/out.mp4' })
    expect(progress[0]).toBeCloseTo(1 / 30)
    expect(progress.at(-1)).toBe(1)
    // the recorder's stop-flush chunk landed in the sink, and only a committed
    // finish followed it
    expect(sinkTrace).toEqual({ writes: [0], finish: [false] })

    // first frame: only slide 0 at full opacity
    expect(state.draws[0]).toEqual({ tag: 's0', alpha: 1 })
    // last frame (frame 29, t=2900ms): only slide 1
    expect(state.draws.at(-1)).toEqual({ tag: 's1', alpha: 1 })
    // fade frames 11..19 draw the pair (frame 10's alpha is still 0 → s0 alone):
    // frame 15 (t=1500ms, alpha 0.5) is the pair at draws 19..20
    expect(state.draws[19]).toEqual({ tag: 's0', alpha: 1 })
    expect(state.draws[20]!.tag).toBe('s1')
    expect(state.draws[20]!.alpha).toBeCloseTo(0.5)
    // every slide was fetched exactly once and released by the end (BUG-1300)
    expect(log.fetched).toEqual(['s0', 's1'])
    expect(log.released.sort()).toEqual(['s0', 's1'])
  })

  it('resolves null for an empty timeline and honors cooperative cancel', async () => {
    const empty = await recordVideoTimeline(
      {
        timeline: { items: [], totalMs: 0 },
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider([], { released: [], fetched: [] }),
        sink: fakeSink({ writes: [], finish: [] }),
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
    const log = { released: [] as string[], fetched: [] as string[] }
    const sinkTrace: SinkTrace = { writes: [], finish: [] }
    const cancel = { current: true } // cancel before the first frame lands
    const cancelled = await recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], log),
        sink: fakeSink(sinkTrace),
        mimeType: 'video/webm',
        cancel,
      },
      fakeHost(state),
    )
    expect(cancelled).toBeNull()
    expect(state.requestedFrames).toBe(0)
    expect(state.recorderCalls).toEqual(['start:video/webm:1000', 'stop'])
    // a cancelled run aborts the stream — the temp file is discarded
    expect(sinkTrace.finish).toEqual([true])
    expect(log.released).toEqual(['s0']) // the pre-loop decode is still freed
  })

  it('fails with the encoder error instead of returning a truncated blob', async () => {
    // BUG-1208: a runtime MediaRecorder error after start must surface as a
    // thrown failure — not a silent "successful" truncated recording
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const log = { released: [] as string[], fetched: [] as string[] }
    const sinkTrace: SinkTrace = { writes: [], finish: [] }
    const p = recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], log),
        sink: fakeSink(sinkTrace),
        mimeType: 'video/mp4',
      },
      fakeHost(state, { errorAfterStart: new Error('encoder exploded') }),
    )
    await expect(p).rejects.toThrow('video recording failed: encoder exploded')
    // the recorder is still torn down, and no frames land after the error
    expect(state.recorderCalls).toEqual(['start:video/mp4:1000', 'stop'])
    expect(state.requestedFrames).toBe(0)
    // the failure aborts the output stream and frees the decoded window
    expect(sinkTrace.finish).toEqual([true])
    expect(log.released).toEqual(['s0'])
  })

  it('fails when stop() never settles (encoder hang) instead of awaiting forever', async () => {
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const sinkTrace: SinkTrace = { writes: [], finish: [] }
    const p = recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], { released: [], fetched: [] }),
        sink: fakeSink(sinkTrace),
        mimeType: 'video/webm',
      },
      fakeHost(state, { hangOnStop: true }),
    )
    // the collapsed fake timers fire the stop-timeout guard on a microtask
    await expect(p).rejects.toThrow('video recorder did not finish (encoder hang)')
    expect(state.recorderCalls).toEqual(['start:video/webm:1000', 'stop'])
    expect(sinkTrace.finish).toEqual([true])
  })

  it('fails when the recorder stops cleanly but produced no data', async () => {
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const sinkTrace: SinkTrace = { writes: [], finish: [] }
    const p = recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], { released: [], fetched: [] }),
        sink: fakeSink(sinkTrace),
        mimeType: 'video/webm',
      },
      fakeHost(state, { noData: true }),
    )
    await expect(p).rejects.toThrow('video recorder produced no data')
    expect(sinkTrace.finish).toEqual([true])
  })

  it('still resolves null on cooperative cancel even with no data', async () => {
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const cancelled = await recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], { released: [], fetched: [] }),
        sink: fakeSink({ writes: [], finish: [] }),
        mimeType: 'video/webm',
        cancel: { current: true },
      },
      fakeHost(state, { noData: true }),
    )
    expect(cancelled).toBeNull() // cancel is not a failure — no error thrown
  })

  it('resolves null when a cancelled run races a hung encoder (BUG-1301)', async () => {
    // a user Cancel that arrives before the encoder wedged must stay the quiet
    // null abort: the stop-timeout rejection fires unconditionally, and before
    // BUG-1301 it surfaced as "export failed" for a user-initiated cancel
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
    }
    const sinkTrace: SinkTrace = { writes: [], finish: [] }
    const cancelled = await recordVideoTimeline(
      {
        timeline: twoSlideTimeline(),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], { released: [], fetched: [] }),
        sink: fakeSink(sinkTrace),
        mimeType: 'video/webm',
        cancel: { current: true },
      },
      fakeHost(state, { hangOnStop: true }),
    )
    expect(cancelled).toBeNull() // cancel wins over the hang — no error thrown
    expect(state.recorderCalls).toEqual(['start:video/webm:1000', 'stop'])
    // the aborted run discards the temp file like every other cancel path
    expect(sinkTrace.finish).toEqual([true])
  })

  it('aspect-fits each slide into the frame when slideSizes is provided', async () => {
    // BUG-1211: a 4:3 slide among 16:9 ones must pillarbox into the 16:9 frame
    // (240x180 centered) instead of stretching to 320x180. Hard cuts (no
    // transitions) keep exactly one draw per frame.
    const timeline = buildVideoTimeline(
      {
        slides: [{}, {}],
        advanceMs: [500, 500],
        transitions: [
          { kind: 'none', durationMs: null },
          { kind: 'none', durationMs: null },
        ],
        options: { fps: 10, useTimings: true, secondsPerSlide: 5, includeTransitions: true },
      },
      never,
    )
    const state: FakeHostState = {
      draws: [],
      requestedFrames: 0,
      chunks: [],
      recorderCalls: [],
      drawBoxes: [],
    }
    const recorded = await recordVideoTimeline(
      {
        timeline,
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: bitmapProvider(['s0', 's1'], { released: [], fetched: [] }),
        slideSizes: [
          { width: 1600, height: 900 }, // 16:9 — fills the frame
          { width: 1280, height: 960 }, // 4:3 — pillarboxed
        ],
        sink: fakeSink({ writes: [], finish: [] }),
        mimeType: 'video/webm',
      },
      fakeHost(state),
    )
    expect(recorded).not.toBeNull()
    expect(state.drawBoxes!.length).toBe(10) // 1s of content at 10fps
    // frames 0..4 hold the 16:9 slide full-frame
    expect(state.drawBoxes![0]).toEqual([0, 0, 320, 180])
    expect(state.drawBoxes![4]).toEqual([0, 0, 320, 180])
    // frames 5..9 hold the 4:3 slide pillarboxed with black side bars
    expect(state.drawBoxes![5]).toEqual([40, 0, 240, 180])
    expect(state.drawBoxes![9]).toEqual([40, 0, 240, 180])
  })
})

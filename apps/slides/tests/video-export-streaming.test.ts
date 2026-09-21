import { describe, expect, it } from 'vitest'
import {
  createSlideBitmapProvider,
  createVideoFileSink,
  recordVideoTimeline,
  type RecorderHost,
  type SlideBitmap,
  type SlideBitmapProvider,
  type VideoChunkSink,
} from '../src/renderer/video-export'
import { buildVideoTimeline } from '../src/renderer/video-plan'

/**
 * BUG-1300 streaming regressions. The pre-streaming pipeline materialized
 * three deck-sized footprints at once: every slide PNG as a base64 string,
 * every slide decoded for the whole recording, and the whole recorded file
 * (chunk array + arrayBuffer copy + structured clone). These tests pin the
 * streaming contracts that keep memory O(1 slide) + O(chunk):
 * - the recording window holds at most the crossfading pair of bitmaps and
 *   releases each one as the timeline moves past it (allocation counter);
 * - the provider fetches each slide exactly once — nothing accumulates
 *   behind it (no string array, no decoded-deck array);
 * - recorder chunks reach the sink in flush order, serialized, and the file
 *   stream commits only after every append resolved;
 * - cancel / provider failure / append failure each abort the stream (the
 *   temp file is discarded) and free the window.
 */

/**
 * Collapsed-timer host whose recorder flushes a dataavailable chunk every
 * `everyFrames` requestFrame()s (the timeslice), plus the final flush on
 * stop. `onFrame` runs at the top of every requestFrame for test hooks.
 */
function chunkingHost(
  everyFrames: number,
  onFrame?: () => void,
): {
  host: RecorderHost
  state: { requestedFrames: number; recorderCalls: string[] }
} {
  const state = { requestedFrames: 0, recorderCalls: [] as string[] }
  let dataHandler: ((e: { data: Blob }) => void) | null = null
  let stopHandler: (() => void) | null = null
  const host: RecorderHost = {
    createCanvas(width, height) {
      const ctx = {
        canvas: { width, height },
        fillStyle: '#000',
        globalAlpha: 1,
        fillRect() {},
        drawImage() {},
      }
      return {
        ctx: ctx as unknown as CanvasRenderingContext2D,
        captureStream() {
          return {
            getVideoTracks: () => [
              {
                requestFrame: () => {
                  onFrame?.()
                  state.requestedFrames++
                  if (state.requestedFrames % everyFrames === 0) {
                    // timeslice flush: the chunk's payload names its frame
                    dataHandler?.({ data: new Blob([new Uint8Array([state.requestedFrames])]) })
                  }
                },
              },
            ],
          }
        },
      }
    },
    createRecorder(_stream, recOpts) {
      return {
        start(timesliceMs?: number) {
          state.recorderCalls.push(`start:${recOpts.mimeType}:${timesliceMs}`)
        },
        stop() {
          state.recorderCalls.push('stop')
          // the final flush just before onstop, as MediaRecorder does
          dataHandler?.({ data: new Blob([new Uint8Array([255])]) })
          queueMicrotask(() => stopHandler?.())
        },
        ondataavailable(handler) {
          dataHandler = handler
        },
        onstop(handler) {
          stopHandler = handler
        },
        onerror() {},
      }
    },
    now: () => 0,
    setTimeout(fn) {
      queueMicrotask(fn)
      return 0
    },
  }
  return { host, state }
}

/** n slides, hard 300ms holds + 300ms fades between them. */
const deckTimeline = (n: number) =>
  buildVideoTimeline(
    {
      slides: Array.from({ length: n }, () => ({})),
      advanceMs: Array.from({ length: n }, () => 300),
      transitions: Array.from({ length: n }, (_, i) =>
        i === 0 ? { kind: 'none', durationMs: null } : { kind: 'fade', durationMs: 300 },
      ),
      options: {
        fps: 10,
        useTimings: true,
        secondsPerSlide: 5,
        includeTransitions: true,
      },
    },
    () => 0,
  )

/** Sink that flattens each chunk to its first byte (the flush's frame stamp). */
function payloadSink(trace: { writes: number[]; finish: boolean[] }): VideoChunkSink {
  return {
    async write(chunk: Blob) {
      trace.writes.push(new Uint8Array(await chunk.arrayBuffer())[0]!)
    },
    async finish(aborted) {
      trace.finish.push(aborted)
      return aborted ? null : '/out/deck.mp4'
    },
  }
}

/** Counting provider: allocation/release accounting is the memory harness. */
function countingProvider(): {
  provider: SlideBitmapProvider
  stats: () => {
    fetched: number[]
    live: number
    maxLive: number
    released: number
    reFetches: number
  }
} {
  const fetched: number[] = []
  let live = 0
  let maxLive = 0
  let released = 0
  let reFetches = 0

  const provider: SlideBitmapProvider = async (itemIndex) => {
    if (fetched.includes(itemIndex)) reFetches++
    fetched.push(itemIndex)
    await Promise.resolve() // decode is async — the prefetch path exercises single-flight
    live++
    maxLive = Math.max(maxLive, live)
    const bitmap: SlideBitmap = {
      image: { tag: `s${itemIndex}` } as unknown as CanvasImageSource,
      release: () => {
        live--
        released++
      },
    }
    return bitmap
  }
  return { provider, stats: () => ({ fetched, live, maxLive, released, reFetches }) }
}

describe('streaming slide window (BUG-1300)', () => {
  it('keeps at most the crossfading pair decoded, releases every bitmap, fetches each slide once', async () => {
    const { provider, stats } = countingProvider()
    const sinkTrace = { writes: [] as number[], finish: [] as boolean[] }
    const { host } = chunkingHost(1_000 /* never mid-run: only the stop flush */)
    const recorded = await recordVideoTimeline(
      {
        timeline: deckTimeline(12),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: provider,
        sink: payloadSink(sinkTrace),
        mimeType: 'video/mp4',
      },
      host,
    )
    expect(recorded).not.toBeNull()
    const s = stats()
    // every slide was rendered/decoded exactly once — nothing deck-sized
    // accumulated behind the provider (no PNG string array, no decoded deck)
    expect(s.reFetches).toBe(0)
    expect(s.fetched).toHaveLength(12)
    // the window never held more than the crossfading pair
    expect(s.maxLive).toBeLessThanOrEqual(2)
    // and everything was released by the end — no leaked bitmaps
    expect(s.live).toBe(0)
    expect(s.released).toBe(12)
  })

  it('frees the window when the provider fails mid-deck and fails the export', async () => {
    const base = countingProvider()
    const provider: SlideBitmapProvider = async (i) => {
      if (i === 3) throw new Error('slide render died')
      return base.provider(i)
    }
    const sinkTrace = { writes: [] as number[], finish: [] as boolean[] }
    const p = recordVideoTimeline(
      {
        timeline: deckTimeline(6),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: provider,
        sink: payloadSink(sinkTrace),
        mimeType: 'video/mp4',
      },
      chunkingHost(1_000).host,
    )
    await expect(p).rejects.toThrow('slide render died')
    expect(base.stats().live).toBe(0) // window released on the way out
    expect(sinkTrace.finish).toEqual([true]) // the temp file is discarded
  })

  it('hands recorder chunks to the sink in flush order and commits after them', async () => {
    const { provider, stats } = countingProvider()
    const sinkTrace = { writes: [] as number[], finish: [] as boolean[] }
    // a chunk every 5 frames: mid-run flushes plus the final stop flush
    const recorded = await recordVideoTimeline(
      {
        timeline: deckTimeline(3),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: provider,
        sink: payloadSink(sinkTrace),
        mimeType: 'video/mp4',
      },
      chunkingHost(5).host,
    )
    expect(stats().live).toBe(0)
    // 2*(300+300)+300ms = 1500ms → 15 frames → flushes at frames 5,10,15, then 255
    expect(sinkTrace.writes).toEqual([5, 10, 15, 255])
    expect(recorded!.bytes).toBe(4) // four one-byte chunks
    // exactly one settle, a commit, after the last append
    expect(sinkTrace.finish).toEqual([false])
  })

  it('aborts the stream on cancel mid-recording (discard, no commit)', async () => {
    const { provider, stats } = countingProvider()
    const sinkTrace = { writes: [] as number[], finish: [] as boolean[] }
    const cancel = { current: false }
    let frames = 0
    // a flush every 3 frames; cancel right after frame 4 lands
    const { host } = chunkingHost(3, () => {
      frames++
      if (frames >= 4) cancel.current = true
    })
    const cancelled = await recordVideoTimeline(
      {
        timeline: deckTimeline(6),
        fps: 10,
        width: 320,
        height: 180,
        slideBitmaps: provider,
        sink: payloadSink(sinkTrace),
        mimeType: 'video/mp4',
        cancel,
      },
      host,
    )
    expect(cancelled).toBeNull()
    expect(stats().live).toBe(0)
    expect(stats().released).toBe(stats().fetched.length)
    // chunks already flushed were appended in order (frame 3's + the stop
    // flush), then the stream aborted — no partial file is ever committed
    expect(sinkTrace.writes).toEqual([3, 255])
    expect(sinkTrace.finish).toEqual([true])
  })
})

describe('createVideoFileSink (chunked IPC write)', () => {
  it('serializes appends: a slow first chunk never lets a later one overtake', async () => {
    const appendOrder: number[] = []
    let releaseFirst!: (r: { ok: boolean }) => void
    const firstAppend = new Promise<{ ok: boolean }>((resolve) => {
      releaseFirst = resolve
    })
    let calls = 0
    const sink = createVideoFileSink(
      (bytes) => {
        calls++
        appendOrder.push(bytes[0]!)
        return calls === 1 ? firstAppend : Promise.resolve({ ok: true })
      },
      (commit) => Promise.resolve(commit ? { ok: true, path: '/out/x.mp4' } : { ok: true }),
    )
    const w1 = sink.write(new Blob([new Uint8Array([1])]))
    const w2 = sink.write(new Blob([new Uint8Array([2])]))
    // let the first chunk reach its (pending) append; the second chunk is not
    // even handed to IPC until the first resolved
    await new Promise((r) => setTimeout(r, 0))
    expect(appendOrder).toEqual([1])
    releaseFirst({ ok: true })
    await Promise.all([w1, w2])
    expect(appendOrder).toEqual([1, 2])
    expect(await sink.finish(false)).toBe('/out/x.mp4')
  })

  it('a failed append poisons the stream: later writes reject fast, finish discards', async () => {
    const appendOrder: number[] = []
    const finishes: boolean[] = []
    const sink = createVideoFileSink(
      (bytes) => {
        appendOrder.push(bytes[0]!)
        return bytes[0] === 1
          ? Promise.resolve({ ok: true })
          : Promise.resolve({ ok: false, error: 'disk full' })
      },
      (commit) => {
        finishes.push(commit)
        return Promise.resolve({ ok: true, path: '/out/x.mp4' })
      },
    )
    await sink.write(new Blob([new Uint8Array([1])]))
    await expect(sink.write(new Blob([new Uint8Array([2])]))).rejects.toThrow('disk full')
    // poisoned: no further IPC traffic, and the settle discards the temp file
    await expect(sink.write(new Blob([new Uint8Array([3])]))).rejects.toThrow('disk full')
    await expect(sink.finish(false)).rejects.toThrow('disk full')
    expect(appendOrder).toEqual([1, 2])
    expect(finishes).toEqual([false])
    // idempotent: a second settle is a quiet no-op (error paths re-settle)
    expect(await sink.finish(true)).toBeNull()
  })
})

describe('createSlideBitmapProvider (render → decode → drop the blob)', () => {
  it('decodes each fetched PNG once and returns a releasable bitmap', async () => {
    const fetched: number[] = []
    const decoded: number[] = []
    const provider = createSlideBitmapProvider(
      async (i) => {
        fetched.push(i)
        return new Blob([new Uint8Array([i])])
      },
      async (png) => {
        const i = new Uint8Array(await png.arrayBuffer())[0]!
        decoded.push(i)
        return {
          image: { tag: i } as unknown as CanvasImageSource,
          release: () => undefined,
        }
      },
    )
    const b0 = await provider(0)
    const b1 = await provider(1)
    expect(fetched).toEqual([0, 1])
    expect(decoded).toEqual([0, 1]) // each blob went straight to decode — nothing cached
    expect((b0.image as unknown as { tag: number }).tag).toBe(0)
    expect((b1.image as unknown as { tag: number }).tag).toBe(1)
    expect(typeof b0.release).toBe('function')
  })
})

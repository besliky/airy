import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  browserRecorderHost,
  isMediaStream,
  recordVideoTimeline,
  type VideoChunkSink,
} from '../src/renderer/video-export'
import { buildVideoTimeline } from '../src/renderer/video-plan'

/**
 * BUG-1620 regression: the browser host used to hand MediaRecorder a
 * `{ getVideoTracks }` wrapper behind a blind `as MediaStream` cast, so the
 * real constructor rejected it ("parameter 1 is not of type 'MediaStream'")
 * and EVERY export failed. The 42 mock-based pipeline tests could not catch
 * it — their fake hosts played along with any shape.
 *
 * These tests therefore exercise the REAL `browserRecorderHost`, stubbing only
 * the platform globals (MediaStream/MediaRecorder/canvas capture). The fake
 * MediaRecorder mirrors the platform contract: it throws Chromium's exact
 * TypeError unless the constructor argument is an `instanceof MediaStream`.
 * If anyone reintroduces a wrapper (or a blind cast) in the host, these fail.
 */

class FakeMediaStream {
  readonly tracks: Array<{ requestFrame?: () => void }>
  constructor(tracks: Array<{ requestFrame?: () => void }> = []) {
    this.tracks = tracks
  }
  getVideoTracks() {
    return this.tracks
  }
}

class FakeMediaRecorder {
  /** Every construction, in order — what the pipeline handed MediaRecorder. */
  static constructed: Array<{ stream: unknown; opts: unknown }> = []
  // event handlers as assignable properties, exactly like the real platform —
  // the host adapter wires itself with `rec.onstop = ...` assignments
  ondataavailable: ((e: { data: Blob }) => void) | null = null
  onstop: (() => void) | null = null
  onerror: ((e: { error?: unknown }) => void) | null = null
  constructor(stream: unknown, opts: unknown) {
    FakeMediaRecorder.constructed.push({ stream, opts })
    if (!(stream instanceof FakeMediaStream)) {
      // the exact rejection the real platform produces for a non-MediaStream
      throw new TypeError(
        "Failed to construct 'MediaRecorder': parameter 1 is not of type 'MediaStream'.",
      )
    }
  }
  start() {
    // a timeslice flush shortly after start, as the real recorder does
    queueMicrotask(() => this.ondataavailable?.({ data: new Blob([new Uint8Array([1, 2, 3])]) }))
  }
  stop() {
    queueMicrotask(() => {
      this.ondataavailable?.({ data: new Blob([new Uint8Array([4])]) })
      this.onstop?.()
    })
  }
}

const fakeCtx = {
  fillStyle: '#000',
  globalAlpha: 1,
  fillRect() {},
  drawImage() {},
} as unknown as CanvasRenderingContext2D

const getContextSpy = vi.fn(() => fakeCtx)
const captureStreamSpy = vi.fn(function (this: HTMLCanvasElement) {
  return currentStream
})
let currentStream: FakeMediaStream = new FakeMediaStream([])

beforeEach(() => {
  currentStream = new FakeMediaStream([])
  FakeMediaRecorder.constructed.length = 0
  captureStreamSpy.mockClear()
  getContextSpy.mockClear()
  vi.stubGlobal('MediaStream', FakeMediaStream)
  vi.stubGlobal('MediaRecorder', FakeMediaRecorder)
  // jsdom has no canvas capture: pin the two members the host touches
  Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
    configurable: true,
    writable: true,
    value: getContextSpy,
  })
  Object.defineProperty(HTMLCanvasElement.prototype, 'captureStream', {
    configurable: true,
    writable: true,
    value: captureStreamSpy,
  })
})

const originalGetContext = Object.getOwnPropertyDescriptor(
  HTMLCanvasElement.prototype,
  'getContext',
)

afterEach(() => {
  vi.unstubAllGlobals()
  if (originalGetContext)
    Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', originalGetContext)
  delete (HTMLCanvasElement.prototype as { captureStream?: unknown }).captureStream
})

describe('browserRecorderHost (BUG-1620 regression)', () => {
  it('captureStream returns the canvas stream object itself, not a wrapper', () => {
    const requestFrame = vi.fn()
    currentStream = new FakeMediaStream([{ requestFrame }])
    const { ctx, captureStream } = browserRecorderHost.createCanvas(320, 160)
    const stream = captureStream(0)
    // capture ran on a real 320x160 canvas element
    expect(captureStreamSpy).toHaveBeenCalledTimes(1)
    expect(captureStreamSpy).toHaveBeenCalledWith(0)
    const canvas = captureStreamSpy.mock.instances[0] as HTMLCanvasElement
    expect(canvas.tagName).toBe('CANVAS')
    expect(canvas.width).toBe(320)
    expect(canvas.height).toBe(160)
    expect(getContextSpy).toHaveBeenCalledWith('2d')
    // the REAL stream — instanceof, not merely structurally shaped
    expect(stream).toBe(currentStream)
    expect(stream instanceof FakeMediaStream).toBe(true)
    const track = stream.getVideoTracks()[0]
    expect(track && 'requestFrame' in track ? track.requestFrame : undefined).toBe(requestFrame)
    expect(ctx).toBe(fakeCtx)
  })

  it('hands the real MediaStream to the MediaRecorder constructor', () => {
    currentStream = new FakeMediaStream([])
    const { captureStream } = browserRecorderHost.createCanvas(320, 160)
    const stream = captureStream(0)
    browserRecorderHost.createRecorder(stream, {
      mimeType: 'video/webm',
      videoBitsPerSecond: 6_000_000,
    })
    expect(FakeMediaRecorder.constructed).toHaveLength(1)
    const { stream: handed, opts } = FakeMediaRecorder.constructed[0]!
    // the very object captureStream() returned reaches the constructor
    expect(handed).toBe(stream)
    expect(handed instanceof FakeMediaStream).toBe(true)
    expect(opts).toEqual({ mimeType: 'video/webm', videoBitsPerSecond: 6_000_000 })
  })

  it('rejects the legacy wrapper shape with a type guard instead of casting', () => {
    // exactly the shape the old host produced (the BUG-1620 lookalike)
    const legacyWrapper = { getVideoTracks: () => [{ requestFrame: () => undefined }] }
    expect(() =>
      browserRecorderHost.createRecorder(legacyWrapper, {
        mimeType: 'video/webm',
        videoBitsPerSecond: 1,
      }),
    ).toThrow(TypeError)
    expect(() =>
      browserRecorderHost.createRecorder(legacyWrapper, {
        mimeType: 'video/webm',
        videoBitsPerSecond: 1,
      }),
    ).toThrow('did not yield a real MediaStream')
    // the (throwing) platform constructor was never reached with the lookalike
    expect(FakeMediaRecorder.constructed).toHaveLength(0)
    // the guard itself: instanceof decides, structural shape does not fool it
    expect(isMediaStream(new FakeMediaStream([]))).toBe(true)
    expect(isMediaStream(legacyWrapper)).toBe(false)
    expect(isMediaStream(null)).toBe(false)
  })

  it('records end-to-end through the real host: recorder receives the captured stream', async () => {
    const requestFrame = vi.fn()
    currentStream = new FakeMediaStream([{ requestFrame }])
    const writes: Blob[] = []
    const sink: VideoChunkSink = {
      async write(chunk) {
        writes.push(chunk)
      },
      async finish(aborted) {
        return aborted ? null : '/tmp/bug1620-smoke.webm'
      },
    }
    const recorded = await recordVideoTimeline(
      {
        timeline: buildVideoTimeline(
          {
            slides: [{}],
            advanceMs: [200],
            transitions: [{ kind: 'none', durationMs: null }],
            options: { fps: 10, useTimings: true, secondsPerSlide: 5, includeTransitions: false },
          },
          () => 0,
        ),
        fps: 10,
        width: 320,
        height: 160,
        slideBitmaps: async () => ({ image: {} as CanvasImageSource, release() {} }),
        mimeType: 'video/webm',
        sink,
      },
      browserRecorderHost, // the real host — only the platform globals stand in
    )
    // captureStream ran on the canvas and its exact return value is what got
    // recorded (on the old code the platform-mirroring constructor throws here)
    expect(captureStreamSpy).toHaveBeenCalledWith(0)
    expect(FakeMediaRecorder.constructed).toHaveLength(1)
    const handed = FakeMediaRecorder.constructed[0]!.stream
    expect(handed).toBe(currentStream)
    expect(handed instanceof FakeMediaStream).toBe(true)
    // frames were pushed through the stream's requestFrame track
    expect(requestFrame).toHaveBeenCalled()
    // the run committed: timeslice + stop-flush chunks reached the sink
    expect(writes.length).toBe(2)
    expect(recorded).toEqual({ bytes: 4, path: '/tmp/bug1620-smoke.webm' })
  })
})

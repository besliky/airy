/**
 * Video export pipeline (File > Export Video): slide bitmaps are decoded from
 * the renderer's offscreen Konva stage ONE at a time, drawn onto a recorder
 * canvas by the pure timeline plan, and recorded through
 * canvas.captureStream(0) + MediaRecorder — the same encoding path the
 * screen-recording insert already uses, so no new dependency.
 *
 * Memory is O(1 slide), not O(deck) (BUG-1300): at most the crossfading pair
 * of bitmaps is decoded at once (a bitmap is released as soon as the timeline
 * moved past its slide), and the recorded container never materializes — the
 * recorder flushes ~1s timeslice chunks straight into the output sink, which
 * appends them to the main process's temp file as they arrive.
 *
 * Encoder: MediaRecorder muxes MP4 (H.264) in Electron's Chromium when
 * proprietary codecs are available (detected at runtime), else falls back to
 * WebM (VP9/VP8). Transitions play as crossfades (see video-plan.ts); the
 * recording is real-time paced because MediaRecorder timestamps frames by
 * wall clock — a 2-minute deck takes ~2 minutes to export.
 */
import type { VideoTimeline } from './video-plan'
import { fitIntoFrame, sampleTimeline, videoFrameCount } from './video-plan'

/** Container the recording lands in (drives the save dialog's file filter). */
export type VideoExportContainer = 'mp4' | 'webm'

export interface RecorderMimeChoice {
  mimeType: string
  container: VideoExportContainer
}

/** MP4 candidates first (Electron ships H.264; Chromium >= 130 muxes mp4), WebM fallback. */
const MP4_MIME_CANDIDATES = [
  'video/mp4;codecs=avc1.640028',
  'video/mp4;codecs=avc1.42E01E',
  'video/mp4',
] as const
const WEBM_MIME_CANDIDATES = [
  'video/webm;codecs=vp9',
  'video/webm;codecs=vp8',
  'video/webm',
] as const

/**
 * Pick the best MediaRecorder mime the environment supports (pure — the
 * isSupported predicate is injected so tests can simulate both worlds).
 * MP4 beats WebM; null when nothing is supported (record button disabled).
 */
export function pickRecorderMime(
  isSupported: (mime: string) => boolean,
): RecorderMimeChoice | null {
  for (const mime of MP4_MIME_CANDIDATES) {
    if (isSupported(mime)) return { mimeType: mime, container: 'mp4' }
  }
  for (const mime of WEBM_MIME_CANDIDATES) {
    if (isSupported(mime)) return { mimeType: mime, container: 'webm' }
  }
  return null
}

/** Encoding bitrate by output height: 1080p ≈ 10 Mbps, 720p ≈ 6 Mbps. */
export function videoBitrate(height: number): number {
  return height >= 1000 ? 10_000_000 : 6_000_000
}

/**
 * One decoded slide bitmap plus an explicit release. The release matters as
 * much as the decode: an `ImageBitmap` is freed deterministically by
 * `close()` the moment its last frame was drawn, instead of lingering until
 * GC gets around to it (BUG-1300 — a 1920x1080 bitmap is ~8.3 MB, and decks
 * of them used to be held for the whole recording).
 */
export interface SlideBitmap {
  image: CanvasImageSource
  /** Free the decoded bitmap immediately (no-op safe to call once). */
  release(): void
}

/**
 * Decode one timeline item's slide bitmap on demand. Items are requested in
 * timeline order; the provider holds NO state between calls, so nothing
 * deck-sized ever accumulates behind it.
 */
export type SlideBitmapProvider = (itemIndex: number) => Promise<SlideBitmap>

/** Decode a rendered slide PNG blob into a closeable bitmap (browser host). */
export async function decodePngBlob(png: Blob): Promise<SlideBitmap> {
  const bitmap = await createImageBitmap(png)
  return { image: bitmap, release: () => bitmap.close() }
}

/**
 * Wire a per-item PNG source into a bitmap provider: fetch the item's PNG
 * (blob storage — no base64 string is ever built), decode it, and let the
 * blob go in the same tick. The returned provider retains nothing: slide
 * bitmaps live exactly as long as the recorder's 2-item window keeps them.
 */
export function createSlideBitmapProvider(
  fetchPng: (itemIndex: number) => Promise<Blob>,
  decode: (png: Blob) => Promise<SlideBitmap> = decodePngBlob,
): SlideBitmapProvider {
  return async (itemIndex) => decode(await fetchPng(itemIndex))
}

/**
 * Destination for the recorded container bytes (BUG-1300): chunks arrive in
 * recorder order as the timeslice flushes them; `finish` settles the output
 * once — committed (resolves the final path) or aborted (discards, resolves
 * null). The recorder never buffers the file itself.
 */
export interface VideoChunkSink {
  /** Accept one recorded chunk; resolves once the bytes are durably accepted. */
  write(chunk: Blob): Promise<void>
  /**
   * Settle the output. `aborted` (cancel/failure) discards everything
   * written; a committing finish resolves the final file path, null when
   * aborted. Idempotent: only the first call settles.
   */
  finish(aborted: boolean): Promise<string | null>
}

/**
 * Sink over the main process's streaming file channels: each chunk is
 * appended through IPC the moment the recorder flushes it, so the peak
 * renderer footprint is a couple of chunk-sized buffers instead of ~3x the
 * whole file (chunk array + arrayBuffer copy + structured clone). Appends are
 * serialized here — IPC round trips may overtake each other, the file's byte
 * order must not. A failed append poisons the stream: later writes reject
 * fast and a `finish` discards the temp file.
 */
export function createVideoFileSink(
  appendChunk: (bytes: Uint8Array) => Promise<{ ok: boolean; error?: string }>,
  finishStream: (commit: boolean) => Promise<{ ok: boolean; path?: string; error?: string }>,
): VideoChunkSink {
  let streamError: unknown = null
  let settled = false
  let tail: Promise<void> = Promise.resolve()
  return {
    write(chunk) {
      if (streamError !== null) return Promise.reject(streamError)
      const write = tail.then(async () => {
        const bytes = new Uint8Array(await chunk.arrayBuffer())
        const r = await appendChunk(bytes)
        if (!r.ok) throw new Error(r.error ?? 'video export stream append failed')
      })
      tail = write.catch((err) => {
        streamError ??= err
      })
      return write
    },
    async finish(aborted) {
      if (settled) return null
      settled = true
      await tail
      const commit = !aborted && streamError === null
      const r = await finishStream(commit)
      if (streamError !== null) throw streamError
      if (!r.ok) throw new Error(r.error ?? 'video export stream finish failed')
      return commit ? (r.path ?? '') : null
    },
  }
}

export interface RecordTimelineOptions {
  timeline: VideoTimeline
  fps: number
  width: number
  height: number
  /**
   * Decode one timeline item's slide on demand (BUG-1300): the recording
   * window keeps at most the crossfading pair alive and releases each bitmap
   * once the timeline moved past it.
   */
  slideBitmaps: SlideBitmapProvider
  /**
   * Source slide size per timeline item (deck order facts, same indexing as
   * slideBitmaps). Provided, each slide aspect-fits into the frame (mixed-size
   * decks letterbox instead of stretching to the frame's shape — BUG-1211);
   * omitted, every image fills the frame exactly as before.
   */
  slideSizes?: ReadonlyArray<{ width: number; height: number }>
  /** Recorded chunks land here in order (BUG-1300: no whole-file blob). */
  sink: VideoChunkSink
  mimeType: string
  /** Defaults to videoBitrate(height) */
  videoBitsPerSecond?: number
  /** Frame progress callback (done, total) */
  onProgress?: (done: number, total: number) => void
  /** Cooperative cancel: checked between frames; a cancelled run resolves null */
  cancel?: { current: boolean }
}

/**
 * Minimal surface of the recorder/canvas APIs the pipeline needs (mockable in tests).
 * onerror is part of the surface so encoder failures can never be silent: a
 * runtime MediaRecorder error (isTypeSupported said yes, the encoder still
 * died — typically an avc1 level vs frame size mismatch) must fail the export
 * instead of yielding a truncated "successful" file or a hung dialog.
 */
export interface RecorderHost {
  createCanvas(
    width: number,
    height: number,
  ): {
    ctx: CanvasRenderingContext2D
    captureStream(frameRequestRate: number): {
      getVideoTracks(): Array<{ requestFrame?: () => void }>
    }
  }
  createRecorder(
    stream: unknown,
    opts: { mimeType: string; videoBitsPerSecond: number },
  ): {
    /** timesliceMs: flush a dataavailable chunk every ~N ms (BUG-1300). */
    start(timesliceMs?: number): void
    stop(): void
    ondataavailable(handler: (e: { data: Blob }) => void): void
    onstop(handler: () => void): void
    onerror(handler: (e: { error?: unknown }) => void): void
  }
  now(): number
  setTimeout(fn: () => void, ms: number): number
}

/** Browser host implementation (tests inject a fake). */
export const browserRecorderHost: RecorderHost = {
  createCanvas(width, height) {
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) throw new Error('2d canvas context unavailable')
    return {
      ctx,
      // CanvasCaptureMediaStreamTrack carries requestFrame() (not on plain MediaStreamTrack)
      captureStream: (rate: number) => {
        const stream = canvas.captureStream(rate)
        return {
          getVideoTracks: () => stream.getVideoTracks() as CanvasCaptureMediaStreamTrack[],
        }
      },
    }
  },
  createRecorder(stream, opts) {
    const rec = new MediaRecorder(stream as MediaStream, opts)
    return {
      start: (timesliceMs?: number) => rec.start(timesliceMs),
      stop: () => rec.stop(),
      ondataavailable: (handler) => {
        rec.ondataavailable = (e) => handler(e)
      },
      onstop: (handler) => {
        rec.onstop = () => handler()
      },
      onerror: (handler) => {
        rec.onerror = (e) => handler(e)
      },
    }
  },
  now: () => performance.now(),
  setTimeout: (fn, ms) => window.setTimeout(fn, ms),
}

/** Draw one output frame: `from` at full opacity, `to` crossfaded over it at alpha. */
function drawFrame(
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number,
  image: (item: number) => CanvasImageSource | undefined,
  from: number,
  to: number,
  alpha: number,
  rects?: ReadonlyArray<{ dx: number; dy: number; dw: number; dh: number }>,
): void {
  ctx.fillStyle = '#000'
  ctx.fillRect(0, 0, width, height)
  const full = { dx: 0, dy: 0, dw: width, dh: height }
  const fromImg = image(from)
  if (fromImg) {
    const r = rects?.[from] ?? full
    ctx.drawImage(fromImg, r.dx, r.dy, r.dw, r.dh)
  }
  if (to !== from && alpha > 0) {
    const toImg = image(to)
    if (toImg) {
      ctx.globalAlpha = Math.min(1, Math.max(0, alpha))
      const r = rects?.[to] ?? full
      ctx.drawImage(toImg, r.dx, r.dy, r.dw, r.dh)
      ctx.globalAlpha = 1
    }
  }
}

/**
 * Fail the export when recorder.stop() does not settle within this window: a
 * crashed encoder that fires neither dataavailable/stop nor error must not
 * hang the export (and the suspended background throttling) forever.
 */
export const RECORDER_STOP_TIMEOUT_MS = 5000

/**
 * MediaRecorder timeslice: the recorder flushes a dataavailable chunk every
 * ~1s (≈1.25 MB at the 1080p bitrate). Without it the whole container sits in
 * one blob until stop — the renderer-side half of the BUG-1300 OOM.
 */
export const RECORDER_TIMESLICE_MS = 1000

/** Successful recording outcome (failures throw, cancels resolve null). */
export interface RecordedVideo {
  /** Total container bytes handed to the sink */
  bytes: number
  /** Committed file path from the sink (null for non-file sinks) */
  path: string | null
}

/**
 * Record the timeline into the sink. Frames are scheduled on the wall clock
 * (start + i*frameMs) and pushed with track.requestFrame() so the container's
 * timestamps match the plan's durations; each chunk the recorder flushes is
 * written to the sink immediately (ordered, serialized). Slide bitmaps are
 * decoded on demand — the crossfading pair at most — and released as the
 * timeline moves on (BUG-1300). Resolves null when the timeline is empty or
 * the run was cancelled. Throws when the encoder fails (onerror), refuses to
 * finish (stop timeout), produces no data, or the sink rejects a chunk — the
 * caller reports the failure instead of writing a truncated file.
 */
export async function recordVideoTimeline(
  opts: RecordTimelineOptions,
  host: RecorderHost = browserRecorderHost,
): Promise<RecordedVideo | null> {
  const { timeline, fps, width, height } = opts
  if (timeline.items.length === 0) return null
  const totalFrames = videoFrameCount(timeline, fps)
  const frameMs = 1000 / fps
  // slide-bitmap window: single-flight decodes land in `live`, and an item is
  // released the moment a frame's `from` moved past it — frames only ever
  // move forward, so a bitmap below `from` can never be drawn again
  const live = new Map<number, SlideBitmap>()
  const pending = new Map<number, Promise<SlideBitmap>>()
  const ensure = (item: number): Promise<SlideBitmap> => {
    let p = pending.get(item)
    if (!p) {
      p = opts.slideBitmaps(item).then((bitmap) => {
        live.set(item, bitmap)
        return bitmap
      })
      pending.set(item, p)
      // a prefetch nobody comes to await (cancel arrived first) must not
      // surface as an unhandled rejection; awaiting it later still throws
      void p.catch(() => undefined)
    }
    return p
  }
  const releasePast = (item: number) => {
    for (const k of [...live.keys()]) {
      if (k < item) {
        live.get(k)!.release()
        live.delete(k)
      }
    }
  }
  // settle the sink on every exit path: cancel/error discard the temp file,
  // success commits it — the sink is idempotent, and a discard failure must
  // not mask the original outcome
  const discard = async () => {
    await opts.sink.finish(true).catch(() => undefined)
  }
  try {
    // decode the first slide before the recorder exists: a render/decode
    // failure must not leave a started recorder (or an empty recording) behind
    await ensure(0)
    const { ctx, captureStream } = host.createCanvas(width, height)
    const stream = captureStream(0)
    const requestFrame = stream.getVideoTracks()[0]?.requestFrame?.bind(stream.getVideoTracks()[0])
    let chunkBytes = 0
    let sinkError: unknown = null
    let sinkTail: Promise<void> = Promise.resolve()
    const recorder = host.createRecorder(stream, {
      mimeType: opts.mimeType,
      videoBitsPerSecond: opts.videoBitsPerSecond ?? videoBitrate(height),
    })
    recorder.ondataavailable((e) => {
      if (!e.data || e.data.size === 0 || sinkError !== null) return
      chunkBytes += e.data.size
      // ordered hand-off: a chunk is written only after its predecessor's
      // append resolved, and no chunk array ever accumulates (BUG-1300)
      const write = sinkTail.then(() => opts.sink.write(e.data))
      sinkTail = write.catch((err) => {
        sinkError ??= err
      })
    })
    // Encoder errors settle the stopped gate AND fail the run afterwards; some
    // Chromium builds fire dataavailable+stop on error (truncated-but-"ok"),
    // others fire nothing at all — the explicit error flag covers both.
    let recordError: unknown = null
    const stopped = new Promise<void>((resolve) => {
      recorder.onstop(() => resolve())
      recorder.onerror((e) => {
        recordError = e.error ?? new Error('MediaRecorder failed')
        resolve()
      })
    })
    recorder.start(RECORDER_TIMESLICE_MS)
    const start = host.now()
    const sleepTo = (atMs: number) =>
      new Promise<void>((resolve) => host.setTimeout(resolve, Math.max(0, atMs - host.now())))
    // per-slide aspect-fit rectangles (BUG-1211): without slideSizes every
    // image fills the frame (the historical stretch behavior)
    const rects = opts.slideSizes?.map((s) => fitIntoFrame(s.width, s.height, width, height))
    let cancelled = false
    for (let i = 0; i < totalFrames; i++) {
      await sleepTo(start + i * frameMs)
      if (opts.cancel?.current) {
        cancelled = true
        break
      }
      if (recordError !== null || sinkError !== null) break // encoder died or the file fell over — stop pushing frames
      const frame = sampleTimeline(timeline, (i * 1000) / fps)
      releasePast(frame.from)
      if (frame.from === frame.to && frame.from + 1 < timeline.items.length) {
        // prefetch during the hold: the next slide renders/decodes while the
        // hold's frames (identical content) tick, so the crossfade starts on
        // time instead of waiting a render-latency for its first frame
        void ensure(frame.from + 1)
      }
      await ensure(frame.from)
      await ensure(frame.to)
      drawFrame(
        ctx,
        width,
        height,
        (item) => live.get(item)?.image,
        frame.from,
        frame.to,
        frame.alpha,
        rects,
      )
      requestFrame?.()
      opts.onProgress?.(i + 1, totalFrames)
    }
    // let the last pushed frame land on the recorder clock before closing
    await sleepTo(host.now() + frameMs)
    try {
      recorder.stop()
    } catch {
      // an errored recorder refuses stop(); the gate below is already settled
      // by onerror — proceed to the error report
    }
    await Promise.race([
      stopped,
      new Promise<void>((_, reject) =>
        host.setTimeout(
          () => reject(new Error('video recorder did not finish (encoder hang)')),
          RECORDER_STOP_TIMEOUT_MS,
        ),
      ),
    ])
    await sinkTail // every flushed chunk is appended before the file settles
    if (cancelled) {
      await discard()
      return null
    }
    if (recordError !== null) {
      await discard()
      throw new Error(
        `video recording failed: ${String((recordError as Error)?.message ?? recordError)}`,
      )
    }
    if (sinkError !== null) {
      await discard()
      throw new Error(
        `video stream write failed: ${String((sinkError as Error)?.message ?? sinkError)}`,
      )
    }
    if (chunkBytes === 0) throw new Error('video recorder produced no data')
    const path = await opts.sink.finish(false)
    return { bytes: chunkBytes, path }
  } catch (err) {
    // early throws (first-slide decode, stop timeout, no data) skipped the
    // settle paths above — the stream must not outlive the run
    await discard()
    throw err
  } finally {
    releasePast(timeline.items.length) // free everything still in the window
  }
}

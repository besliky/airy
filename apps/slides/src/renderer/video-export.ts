/**
 * Video export pipeline (File > Export Video): slide PNGs drawn onto an
 * offscreen canvas by the pure timeline plan, recorded through
 * canvas.captureStream(0) + MediaRecorder — the same encoding path the
 * screen-recording insert already uses, so no new dependency.
 *
 * Encoder: MediaRecorder muxes MP4 (H.264) in Electron's Chromium when
 * proprietary codecs are available (detected at runtime), else falls back to
 * WebM (VP9/VP8). Transitions play as crossfades (see video-plan.ts); the
 * recording is real-time paced because MediaRecorder timestamps frames by
 * wall clock — a 2-minute deck takes ~2 minutes to export.
 */
import type { VideoTimeline } from './video-plan'
import { sampleTimeline, videoFrameCount } from './video-plan'

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

/** Decode the exported slide PNGs into drawable images, timeline-item order. */
export async function decodePngImages(
  pngsBase64: ReadonlyArray<string>,
): Promise<HTMLImageElement[]> {
  return Promise.all(
    pngsBase64.map(
      (b64) =>
        new Promise<HTMLImageElement>((resolve, reject) => {
          const img = new Image()
          img.onload = () => resolve(img)
          img.onerror = () => reject(new Error('slide bitmap decode failed'))
          img.src = `data:image/png;base64,${b64}`
        }),
    ),
  )
}

export interface RecordTimelineOptions {
  timeline: VideoTimeline
  fps: number
  width: number
  height: number
  /** Slide bitmap per timeline item (same order as timeline.items) */
  slideImages: ReadonlyArray<CanvasImageSource>
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
    start(): void
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
      start: () => rec.start(),
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
  images: ReadonlyArray<CanvasImageSource>,
  from: number,
  to: number,
  alpha: number,
): void {
  ctx.fillStyle = '#000'
  ctx.fillRect(0, 0, width, height)
  const fromImg = images[from]
  if (fromImg) ctx.drawImage(fromImg, 0, 0, width, height)
  if (to !== from && alpha > 0) {
    const toImg = images[to]
    if (toImg) {
      ctx.globalAlpha = Math.min(1, Math.max(0, alpha))
      ctx.drawImage(toImg, 0, 0, width, height)
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
 * Record the timeline into a video Blob. Frames are scheduled on the wall
 * clock (start + i*frameMs) and pushed with track.requestFrame() so the
 * container's timestamps match the plan's durations. Resolves null when the
 * timeline is empty or the run was cancelled. Throws when the encoder fails
 * (onerror), refuses to finish (stop timeout) or produces no data at all —
 * the caller reports the failure instead of writing a truncated file.
 */
export async function recordVideoTimeline(
  opts: RecordTimelineOptions,
  host: RecorderHost = browserRecorderHost,
): Promise<Blob | null> {
  const { timeline, fps, width, height } = opts
  if (timeline.items.length === 0 || opts.slideImages.length === 0) return null
  const totalFrames = videoFrameCount(timeline, fps)
  const frameMs = 1000 / fps
  const { ctx, captureStream } = host.createCanvas(width, height)
  const stream = captureStream(0)
  const requestFrame = stream.getVideoTracks()[0]?.requestFrame?.bind(stream.getVideoTracks()[0])
  const chunks: Blob[] = []
  const recorder = host.createRecorder(stream, {
    mimeType: opts.mimeType,
    videoBitsPerSecond: opts.videoBitsPerSecond ?? videoBitrate(height),
  })
  recorder.ondataavailable((e) => {
    if (e.data && e.data.size > 0) chunks.push(e.data)
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
  recorder.start()

  const start = host.now()
  const sleepTo = (atMs: number) =>
    new Promise<void>((resolve) => host.setTimeout(resolve, Math.max(0, atMs - host.now())))
  let cancelled = false
  for (let i = 0; i < totalFrames; i++) {
    await sleepTo(start + i * frameMs)
    if (opts.cancel?.current) {
      cancelled = true
      break
    }
    if (recordError !== null) break // encoder died — stop pushing frames
    const frame = sampleTimeline(timeline, (i * 1000) / fps)
    drawFrame(ctx, width, height, opts.slideImages, frame.from, frame.to, frame.alpha)
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
  if (recordError !== null)
    throw new Error(
      `video recording failed: ${String((recordError as Error)?.message ?? recordError)}`,
    )
  if (!cancelled && chunks.length === 0) throw new Error('video recorder produced no data')
  return cancelled ? null : new Blob(chunks, { type: opts.mimeType })
}

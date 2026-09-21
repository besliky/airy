/**
 * File actions extracted from App.tsx: save / save-as, image & PDF
 * export. Each function takes the ActionCtx built fresh per call.
 * (Printing lives in components/PrintDialog.tsx — preview + options dialog.)
 */
import type { RenderSlide } from '@airy-office/pptx-render'
import type { ActionCtx } from './action-context'
import { createSlidePngRenderer, renderSlidesToPngBase64 } from './export-render'
import { renderSlideSvg } from './slide-svg'
import { buildVideoTimeline, videoFrameDimensions } from './video-plan'
import {
  createSlideBitmapProvider,
  createVideoFileSink,
  pickRecorderMime,
  recordVideoTimeline,
} from './video-export'
import { t } from './i18n/locale'
import { showToast } from '@airy-office/ui/toast-bus'

/**
 * If a text box/table is still being edited on ⌘S/close-save, blur first so the
 * overlay commits (blur→commitEdit), and save only after the commit lands —
 * otherwise we save pre-edit content and get a dirty-close prompt again.
 */
export async function flushActiveEdit(ctx: ActionCtx): Promise<void> {
  const active = document.activeElement as HTMLElement | null
  if (!active?.isContentEditable) return
  active.blur()
  for (let i = 0; i < 40 && ctx.editingActiveRef.current; i++) {
    await new Promise((r) => setTimeout(r, 50))
  }
}

/**
 * After save, the main process reopens the file with all-new element ids: swap
 * the render tree, mapping selection/edit state to new ids by per-page node ordinal.
 */
export function adoptSavedSlides(ctx: ActionCtx, next: RenderSlide[]): void {
  const remap = (id: string) => {
    const i = ctx.slides[ctx.current]?.nodes.findIndex((n) => n.sourceId === id) ?? -1
    return next[ctx.current]?.nodes[i]?.sourceId ?? null
  }
  ctx.setSelectedIds((ids) => ids.map(remap).filter((x): x is string => x !== null))
  ctx.setEnteredGroupId(null) // Group children ids can't be mapped by top-level ordinal; exit in-group editing after save
  ctx.setEditing((e) => (e && remap(e.sourceId) ? { sourceId: remap(e.sourceId)! } : e))
  ctx.setEditingCell((c) => (c && remap(c.sourceId) ? { ...c, sourceId: remap(c.sourceId)! } : c))
  ctx.setSlides(next)
}

/**
 * Serializes save passes: a call that arrives while a save is in flight waits
 * for it instead of running concurrently. Two overlapping saves write the
 * same file with two `createWriteStream` pipes — interleaved zip streams,
 * truncated pptx, or EPERM/EBUSY on Windows. The queue is a simple promise
 * chain: each caller awaits the previous tail, then runs its own pass.
 */
let saveTail: Promise<unknown> | null = null

/**
 * Runs `pass` after the in-flight save (if any) finishes, and becomes the
 * tail subsequent saves wait on. A pass that throws still releases the
 * queue; the error propagates to its own caller only.
 */
async function runSerialized<T>(pass: () => Promise<T>): Promise<T> {
  const prior = saveTail
  const current = (async (): Promise<T> => {
    if (prior) await prior.catch(() => undefined)
    return pass()
  })()
  const tail = current.then(
    () => undefined,
    () => undefined,
  )
  saveTail = tail
  void current.then(
    () => {
      if (saveTail === tail) saveTail = null
    },
    () => {
      if (saveTail === tail) saveTail = null
    },
  )
  return current
}

export async function save(getCtx: () => ActionCtx, quiet = false): Promise<boolean> {
  return runSerialized(async () => {
    // resolved only now: a queued pass must remap selection against the tree the prior save adopted
    const ctx = getCtx()
    await flushActiveEdit(ctx)
    await ctx.flushNotes()
    const r = await window.slidesApi.save()
    if (r.ok) {
      if (r.slides) adoptSavedSlides(ctx, r.slides)
      if (r.path) ctx.setPath(r.path)
      ctx.setDirty(false)
      const saved = t('appStatusSaved')
      ctx.setStatus(saved)
      if (!quiet) showToast(saved)
    } else {
      const failed = t('appStatusSaveFailed', { error: r.error ?? t('appErrorCanceled') })
      ctx.setStatus(failed)
      // quiet suppresses the success toast only — a failed save (incl. the 30s
      // auto-save) must surface, or edits silently stop reaching disk
      showToast(failed, 'error')
    }
    return r.ok
  })
}

export async function saveAs(getCtx: () => ActionCtx): Promise<void> {
  // Same queue as save(): Save + Save As (or double Save As) write through
  // the same main-process pipe and would interleave without it.
  await runSerialized(async () => {
    const ctx = getCtx()
    await flushActiveEdit(ctx)
    await ctx.flushNotes()
    const name = ctx.path?.split('/').pop() ?? 'presentation.pptx'
    const r = await window.slidesApi.saveAs(name)
    if (r.ok) {
      if (r.slides) adoptSavedSlides(ctx, r.slides)
      ctx.setPath(r.path ?? ctx.path)
      ctx.setDirty(false)
      const saved = t('appStatusSavedAs')
      ctx.setStatus(saved)
      showToast(saved)
    } else if (r.error) {
      // a canceled dialog returns ok:false without error — only real write
      // failures surface, matching the docs/sheets save-as feedback
      const failed = t('appStatusSaveFailed', { error: r.error })
      ctx.setStatus(failed)
      showToast(failed, 'error')
    }
  })
}

/** Export base name: file name without the .pptx extension */
export function exportBaseName(ctx: ActionCtx): string {
  return (ctx.path?.split('/').pop() ?? t('appUntitledPresentation')).replace(/\.pptx$/i, '')
}

/** Export as images: each page (skipping hidden ones) rendered offscreen to 2x PNG, written to disk by the main process */
export async function exportImages(ctx: ActionCtx): Promise<void> {
  const visible = ctx.slides.filter((s) => !s.hidden)
  if (visible.length === 0) {
    ctx.setStatus(t('appExportNoSlides'))
    return
  }
  const dir = await window.slidesApi.pickExportDir()
  if (!dir) return
  ctx.setStatus(t('appExportImagesProgress', { count: visible.length }))
  try {
    const pngs = await renderSlidesToPngBase64(visible, ctx.images)
    const r = await window.slidesApi.exportImages({
      dir,
      baseName: exportBaseName(ctx),
      pngsBase64: pngs,
    })
    ctx.setStatus(
      r.ok
        ? t('appExportImagesDone', { count: r.paths?.length ?? 0, dir })
        : t('appExportImagesFailed', { error: r.error ?? t('appUnknownError') }),
    )
  } catch (err) {
    ctx.setStatus(t('appExportImagesFailed', { error: String(err) }))
  }
}

/** Export PDF page layouts (the print sheet's subset that makes sense as a file) */
export type PdfExportLayout = 'full' | 'notes' | 'handout2' | 'handout3'

/**
 * Export as PDF: each page painted as an inline SVG through the same
 * printToPDF-over-DOM pipeline the print sheet uses — the PDF carries real,
 * selectable text. A slide whose SVG assembly fails falls back to the raster
 * page (2x PNG) individually; the rest of the deck stays vector. Non-'full'
 * layouts reuse the print sheet's page assembly (notes pages / handouts).
 */
export async function exportPdf(ctx: ActionCtx, layout: PdfExportLayout = 'full'): Promise<void> {
  const visible = ctx.slides
    .map((slide, deckIndex) => ({ slide, deckIndex }))
    .filter((v) => !v.slide.hidden)
  if (visible.length === 0) {
    ctx.setStatus(t('appExportNoSlides'))
    return
  }
  const target = await window.slidesApi.pickExportPdfPath(`${exportBaseName(ctx)}.pdf`)
  if (!target) return
  ctx.setStatus(t('appExportPdfProgress'))
  try {
    const pages: Array<{ svg?: string; pngBase64?: string }> = []
    for (const { slide } of visible) {
      try {
        // per-slide id prefix: every page inlines into ONE print document and
        // url(#id) is document-global (BUG-1104) — 'p0-grad0' cannot collide
        // with 'p1-grad0'
        pages.push({ svg: renderSlideSvg(slide, ctx.images, `p${pages.length}-`) })
      } catch {
        // raster fallback for this slide only (unexpected node structures)
        const [png] = await renderSlidesToPngBase64([slide], ctx.images)
        if (!png) throw new Error('slide render failed')
        pages.push({ pngBase64: png })
      }
    }
    // notes ride along only for the notes layout (fetched by deck position)
    const notes =
      layout === 'notes'
        ? await Promise.all(visible.map((v) => window.slidesApi.getNotes(v.deckIndex)))
        : undefined
    const r = await window.slidesApi.exportPdf({
      filePath: target,
      pages,
      widthPx: visible[0].slide.widthPx,
      heightPx: visible[0].slide.heightPx,
      layout,
      ...(notes ? { notes } : {}),
    })
    ctx.setStatus(
      r.ok
        ? t('appExportPdfDone', { path: r.path ?? '' })
        : t('appExportPdfFailed', { error: r.error ?? t('appUnknownError') }),
    )
  } catch (err) {
    ctx.setStatus(t('appExportPdfFailed', { error: String(err) }))
  }
}

// ── Export video (File > Export Video) ────────────────────────────────────────

/** Dialog-facing video export settings (resolution preset names the slide height). */
export interface VideoExportSettings {
  fps: number
  heightPreset: 720 | 1080
  /** Pace slides by rehearsed auto-advance times where recorded */
  useTimings: boolean
  /** Dwell for slides without timings (seconds), or all slides when timings are off */
  secondsPerSlide: number
  /** Play transitions as crossfades; false = hard cuts */
  includeTransitions: boolean
}

/** Coarse pipeline phase for the progress callback (the dialog renders both). */
export type VideoExportPhase = 'render' | 'record'

/**
 * Terminal outcome for the video-export dialog. Failures carry a localized
 * reason for the dialog's in-body alert line (UX-1205: the status bar sits
 * behind the modal's dimming, invisible until the dialog closes); user
 * aborts (save-dialog dismiss, Cancel) report no reason — they are not
 * errors.
 */
export interface VideoExportOutcome {
  ok: boolean
  /** Localized failure reason (absent on success and user aborts) */
  error?: string
}

/**
 * Export the deck as a video: an offscreen canvas records the slides through
 * MediaRecorder (mp4 when the Chromium build muxes it, else WebM), paced by
 * the pure timeline (rehearsed timings or the fallback dwell) with
 * transitions as crossfades.
 *
 * The pipeline streams end to end (BUG-1300): each slide is rendered to a PNG
 * blob and decoded only when the recording window reaches it (the crossfading
 * pair at most, released as the timeline moves on), and every recorder chunk
 * is appended to the main process's temp file as it is flushed — the peak
 * footprint is O(1 slide) + a couple of chunks instead of O(deck PNGs +
 * decoded deck + 3x the recorded file), which OOMed the renderer on long
 * decks. Recording is real-time paced — MediaRecorder timestamps frames by
 * the wall clock — so export duration ≈ video duration. Failures return a
 * localized reason for the dialog's alert line; user aborts return ok:false
 * without one.
 */
export async function exportVideo(
  ctx: ActionCtx,
  settings: VideoExportSettings,
  onProgress?: (phase: VideoExportPhase, done: number, total: number) => void,
  cancel?: { current: boolean },
): Promise<VideoExportOutcome> {
  const visible = ctx.slides.filter((s) => !s.hidden)
  if (visible.length === 0) {
    ctx.setStatus(t('appExportNoSlides'))
    return { ok: false, error: t('appExportNoSlides') }
  }
  if (typeof MediaRecorder === 'undefined') {
    ctx.setStatus(t('appExportVideoFailed', { error: t('appExportVideoNoEncoder') }))
    return { ok: false, error: t('appExportVideoNoEncoder') }
  }
  const mime = pickRecorderMime((m) => MediaRecorder.isTypeSupported(m))
  if (!mime) {
    ctx.setStatus(t('appExportVideoFailed', { error: t('appExportVideoNoEncoder') }))
    return { ok: false, error: t('appExportVideoNoEncoder') }
  }
  const first = visible[0]!
  const dims = videoFrameDimensions(first.widthPx, first.heightPx, settings.heightPreset)
  // deck facts for the timeline: per-slide transitions + rehearsed timings
  const [transitions, advanceMs] = await Promise.all([
    Promise.all(ctx.slides.map((_, i) => window.slidesApi.getTransition(i))),
    window.slidesApi.getAdvanceTimes(),
  ])
  const timeline = buildVideoTimeline({
    slides: ctx.slides,
    advanceMs,
    transitions,
    options: settings,
  })
  const target = await window.slidesApi.pickExportVideoPath(
    `${exportBaseName(ctx)}.${mime.container}`,
    mime.container,
  )
  if (!target) return { ok: false } // save dialog dismissed — back to options, no error
  // recording is wall-clock paced: suspend background timer throttling for the
  // run (a minimized window would otherwise clamp frame timers to 1s)
  await window.slidesApi.setVideoExportActive(true)
  // on-demand slide renderer: mounted before anything can fail mid-run and
  // torn down on every exit below
  const renderer = createSlidePngRenderer(visible, ctx.images, dims.width / first.widthPx)
  try {
    // streaming output (BUG-1300): the temp file is opened up front and every
    // recorder chunk is appended as flushed — the container never exists as a
    // renderer-side blob. Commit is the atomic last step, so a cancelled or
    // failed run leaves no partial file behind (only a dot-prefixed temp,
    // which the abort discards).
    const begin = await window.slidesApi.beginVideoFileStream(target)
    if (!begin.ok || begin.token === undefined) {
      throw new Error(begin.error ?? t('appUnknownError'))
    }
    const sink = createVideoFileSink(
      (bytes) => window.slidesApi.appendVideoFileStream(begin.token!, bytes),
      (commit) => window.slidesApi.finishVideoFileStream(begin.token!, commit),
    )
    const recorded = await recordVideoTimeline({
      timeline,
      fps: settings.fps,
      width: dims.width,
      height: dims.height,
      slideBitmaps: createSlideBitmapProvider(renderer.renderPng),
      slideSizes: visible.map((s) => ({ width: s.widthPx, height: s.heightPx })),
      mimeType: mime.mimeType,
      sink,
      onProgress: (done, total) => onProgress?.('record', done, total),
      ...(cancel ? { cancel } : {}),
    })
    if (!recorded) return { ok: false } // cancelled (user abort) — temp discarded in the pipeline
    ctx.setStatus(t('appExportVideoDone', { path: recorded.path ?? target }))
    return { ok: true }
  } catch (err) {
    const error = String(err)
    ctx.setStatus(t('appExportVideoFailed', { error }))
    return { ok: false, error }
  } finally {
    renderer.dispose()
    await window.slidesApi.setVideoExportActive(false)
  }
}

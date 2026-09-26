/**
 * Slide-show tab actions extracted from App.tsx: starting shows,
 * presenter view, custom shows, rehearsal timings, and hiding slides.
 * Functions take the ActionCtx built fresh per call.
 */
import type { ActionCtx } from './action-context'
import type { CustomShow } from './slideshow-utils'
import { t } from './i18n/locale'

/** Instant black curtain under the upcoming show (body::after overlay): painted on the
 *  very next frame after the click, it hides the React mount + window-snap latency.
 *  SlideShowView lifts it once the show is revealed (and on unmount, for early exits). */
export function dropShowCurtain(): void {
  document.body.classList.add('show-curtain')
}
export function liftShowCurtain(): void {
  document.body.classList.remove('show-curtain')
}

export function startSlideShow(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  // From start: jump to the first unhidden slide (if all hidden, still start at slide 1)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: fromStart ? Math.max(0, first) : ctx.current })
}

export function exitSlideShow(ctx: ActionCtx, lastIndex: number): void {
  ctx.setSlideShow(null)
  ctx.setCurrent(lastIndex)
}

/** Update the custom show list overwrite-style and persist */
export function updateCustomShows(ctx: ActionCtx, shows: CustomShow[]): void {
  ctx.setCustomShows(shows)
  if (ctx.path) {
    try {
      localStorage.setItem(`ai-slides-custom-shows:${ctx.path}`, JSON.stringify(shows))
    } catch {
      // Degrade to memory-only when localStorage is full/unavailable
    }
  }
}

/** Play a custom show: only its included slides (filtering out-of-range indexes stale after deletions) */
export function playCustomShow(ctx: ActionCtx, show: CustomShow): void {
  if (ctx.slideShow || ctx.presenter) return
  const order = show.slideIndices.filter((i) => i >= 0 && i < ctx.slides.length)
  if (order.length === 0) {
    ctx.setStatus(t('appStatusCustomShowEmpty'))
    return
  }
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  ctx.setCustomShowDlgOpen(false)
  ctx.setSlideShow({ startAt: order[0]!, customOrder: order })
}

/** Rehearsal timing: show from the start and record each slide's dwell time */
export function startRehearseShow(ctx: ActionCtx): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: Math.max(0, first), rehearse: true })
}

/**
 * Record Slide Show (PAR-314): run the show from the start with an explicit
 * recording session — pause/resume/stop HUD on top of the rehearsal clock.
 * Narration is not recorded (headless/CI has no audio input; follow-up).
 */
export function startRecordShow(ctx: ActionCtx): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setSlideShow({ startAt: Math.max(0, first), rehearse: true, record: true })
}

/** Recording/rehearsal ended: stash per-slide dwell milliseconds; after exiting the show, prompt "save?" */
export function onRehearseDone(ctx: ActionCtx, perPageMs: number[]): void {
  if (perPageMs.some((ms) => ms > 0))
    ctx.setPendingRehearse({ ms: perPageMs, record: ctx.slideShow?.record === true })
}

/** Save rehearsal/recorded timings: write each slide's dwell as its auto-advance time (<p:transition advTm>, exact milliseconds — UX-1768) */
export async function saveRehearseTimings(ctx: ActionCtx): Promise<void> {
  if (!ctx.pendingRehearse) return
  const times = ctx.pendingRehearse.ms
    .map((ms, i) => ({ slideIndex: i, ms }))
    .filter((t) => t.ms > 0)
  ctx.setPendingRehearse(null)
  const ok = await window.slidesApi.setAdvanceTimes({ times })
  if (ok) {
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusRehearseSaved', { count: times.length }))
  }
}

/** Presenter view (single-window version, entry aligned with the show) */
export function startPresenterView(ctx: ActionCtx, fromStart: boolean): void {
  if (ctx.slides.length === 0 || ctx.slideShow || ctx.presenter) return
  dropShowCurtain()
  ctx.setEditing(null)
  ctx.setCtxMenu(null)
  const first = ctx.slides.findIndex((s) => !s.hidden)
  ctx.setPresenter({ startAt: fromStart ? Math.max(0, first) : ctx.current })
}

export function exitPresenterView(ctx: ActionCtx, lastIndex: number): void {
  ctx.setPresenter(null)
  ctx.setCurrent(lastIndex)
}

/** "Switch to normal show" inside presenter view: seamlessly turns into a full-screen show in this window */
export function switchPresenterToShow(ctx: ActionCtx, lastIndex: number): void {
  ctx.setPresenter(null)
  ctx.setCurrent(lastIndex)
  ctx.setSlideShow({ startAt: lastIndex })
}

export async function toggleHidden(ctx: ActionCtx, index: number): Promise<void> {
  const s = ctx.slides[index]
  if (!s) return
  const updated = await window.slidesApi.setSlideHidden({
    slideIndex: index,
    hidden: !s.hidden,
  })
  if (updated) {
    ctx.applySlide(index, updated)
    ctx.setStatus(
      updated.hidden
        ? t('appStatusSlideHidden', { page: index + 1 })
        : t('appStatusSlideUnhidden', { page: index + 1 }),
    )
  }
}

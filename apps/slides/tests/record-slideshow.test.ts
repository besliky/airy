/**
 * Record Slide Show (PAR-314) regression tests:
 * - Record session model: dwell accumulation per slide, pause/resume freezing the
 *   clock, stop converting to per-slide seconds (same contract as rehearsal)
 * - Timing round-trip: recorded auto-advance times (advTm) survive a full
 *   engine save + reopen, and the slide XML carries the advTm attribute that
 *   PowerPoint / python-pptx read as the auto-advance timing
 * - UI state wiring: the ribbon entry starts a recording show, the recorded
 *   dwell is tagged as coming from Record Slide Show, and saving writes
 *   milliseconds through the existing setAdvanceTimes IPC
 */
import { describe, expect, it, vi } from 'vitest'

import {
  currentRecordMs,
  finishRecord,
  formatClock,
  pauseRecord,
  resumeRecord,
  startRecord,
  switchRecordPage,
  totalRecordMs,
} from '../src/renderer/slideshow-utils'
import {
  createBlankPptx,
  getSlideAdvanceTime,
  insertBlankSlide,
  openPptx,
  savePptx,
  setSlideAdvanceTime,
} from '@airy-office/pptx-engine'
import * as showActions from '../src/renderer/show-actions'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/i18n/locale', () => ({ t: (k: string) => k }))

describe('record session accumulation', () => {
  it('accumulates dwell per slide while recording; revisits keep adding', () => {
    let t = startRecord(3, 0, 1000)
    t = switchRecordPage(t, 1, 3500) // slide 0 dwelled 2.5s
    expect(t.perPageMs).toEqual([2500, 0, 0])
    t = switchRecordPage(t, 0, 4000) // slide 1 dwelled 0.5s, back to slide 0
    expect(t.perPageMs).toEqual([2500, 500, 0])
    expect(finishRecord(t, 6000)).toEqual([5, 1, 0]) // 4.5s rounds to 5; 0.5s counts at least 1
  })

  it('pausing banks the current dwell and freezes both clocks; resume starts a fresh window', () => {
    let t = startRecord(2, 0, 1000)
    t = pauseRecord(t, 3000) // slide 0 banked 2s
    expect(t.phase).toBe('paused')
    expect(currentRecordMs(t, 10_000)).toBe(2000) // frozen while paused
    expect(totalRecordMs(t, 10_000)).toBe(2000)
    t = resumeRecord(t, 5000)
    expect(t.phase).toBe('recording')
    // current-slide clock = banked dwell + the fresh post-resume window (revisits keep adding)
    expect(currentRecordMs(t, 6000)).toBe(3000)
    expect(totalRecordMs(t, 6000)).toBe(3000)
  })

  it('page turns while paused move on without accumulating paused time', () => {
    let t = startRecord(3, 0, 0)
    t = pauseRecord(t, 2000)
    t = switchRecordPage(t, 1, 10_000) // 8s paused dwell must not count
    expect(t.perPageMs).toEqual([2000, 0, 0])
    expect(t.currentIndex).toBe(1)
    t = resumeRecord(t, 11_000)
    t = switchRecordPage(t, 2, 13_500)
    expect(finishRecord(t, 13_500)).toEqual([2, 3, 0]) // 2.5s recorded on slide 1 (rounds up)
  })

  it('double pause / double resume are no-ops', () => {
    let t = startRecord(1, 0, 0)
    t = pauseRecord(t, 1000)
    const again = pauseRecord(t, 5000)
    expect(again).toEqual(t)
    t = resumeRecord(t, 6000)
    expect(resumeRecord(t, 7000)).toEqual(t)
  })

  it('a clock going backwards never produces negatives', () => {
    const t = startRecord(1, 0, 5000)
    expect(finishRecord(t, 4000)).toEqual([0])
    expect(currentRecordMs(t, 4000)).toBe(0)
  })

  it('HUD clocks and the rehearsal clock share the same m:ss format', () => {
    expect(formatClock(65_400)).toBe('1:05')
    expect(formatClock(-1)).toBe('0:00')
  })
})

describe('recorded timing round-trip through the pptx file', () => {
  it('advTm written by recording survives save + reopen and lands in the slide XML', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(insertBlankSlide(opened, 0)).not.toBeNull()
    setSlideAdvanceTime(opened.deck.slides[0]!, 5000)
    setSlideAdvanceTime(opened.deck.slides[1]!, null) // explicit clear on the second slide

    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideAdvanceTime(reopened.deck.slides[0]!)).toBe(5000)
    expect(getSlideAdvanceTime(reopened.deck.slides[1]!)).toBeNull()
    // The structural artifact PowerPoint / python-pptx read: advTm on <p:transition>
    expect(reopened.deck.slides[0]!.bodySuffix).toContain('advTm="5000"')
  })
})

// ── UI state wiring (ribbon entry → recording show → save prompt) ────────────────

function makeCtx(over: Record<string, unknown>): ActionCtx {
  return {
    slides: [{}, {}, {}],
    slideShow: null,
    setEditing: vi.fn(),
    setCtxMenu: vi.fn(),
    setSlideShow: vi.fn(),
    setPendingRehearse: vi.fn(),
    setDirty: vi.fn(),
    setStatus: vi.fn(),
    ...over,
  } as unknown as ActionCtx
}

describe('record slide show UI states', () => {
  it('the ribbon Record entry starts a show with rehearse+record from the first unhidden slide', () => {
    const ctx = makeCtx({ slides: [{ hidden: true }, {}] })
    showActions.startRecordShow(ctx)
    expect(ctx.setSlideShow).toHaveBeenCalledWith({ startAt: 1, rehearse: true, record: true })
  })

  it('a normal show never gets the record flag', () => {
    const ctx = makeCtx({})
    showActions.startRehearseShow(ctx)
    expect(ctx.setSlideShow).toHaveBeenCalledWith({ startAt: 0, rehearse: true })
  })

  it('finishing in a recording show tags the pending timings as recorded', () => {
    const ctx = makeCtx({ slideShow: { startAt: 0, rehearse: true, record: true } })
    showActions.onRehearseDone(ctx, [3, 0, 2])
    expect(ctx.setPendingRehearse).toHaveBeenCalledWith({ sec: [3, 0, 2], record: true })
  })

  it('finishing a plain rehearsal keeps the rehearse tagging', () => {
    const ctx = makeCtx({ slideShow: { startAt: 0, rehearse: true } })
    showActions.onRehearseDone(ctx, [3, 0, 2])
    expect(ctx.setPendingRehearse).toHaveBeenCalledWith({ sec: [3, 0, 2], record: false })
  })

  it('all-zero dwell (never shown) does not open the save prompt', () => {
    const ctx = makeCtx({ slideShow: { startAt: 0, record: true } })
    showActions.onRehearseDone(ctx, [0, 0, 0])
    expect(ctx.setPendingRehearse).not.toHaveBeenCalled()
  })

  it('saving writes seconds as milliseconds through setAdvanceTimes and clears the prompt', async () => {
    const setAdvanceTimes = vi.fn(async () => true)
    const ctx = makeCtx({ pendingRehearse: { sec: [2, 0, 5], record: true } })
    vi.stubGlobal('window', { slidesApi: { setAdvanceTimes } })
    await showActions.saveRehearseTimings(ctx)
    expect(setAdvanceTimes).toHaveBeenCalledWith({
      times: [
        { slideIndex: 0, ms: 2000 },
        { slideIndex: 2, ms: 5000 },
      ],
    })
    expect(ctx.setPendingRehearse).toHaveBeenCalledWith(null)
    expect(ctx.setDirty).toHaveBeenCalledWith(true)
    vi.unstubAllGlobals()
  })
})

/**
 * UX-1704: split-view scroll-sync. The proportion math and the loop
 * protection are pure (preview/scroll-sync.ts) and behavior-tested here; the
 * wiring — CM scroll events → rAF coalescing → gx:scrollTo/gx:scrolled — is
 * pinned by source assertions in editor-prefs-wiring.test.ts, because jsdom
 * has no layout and cannot produce real scroll geometry.
 */
import { describe, expect, it } from 'vitest'

import {
  ScrollSyncGate,
  SYNC_ECHO_WINDOW_MS,
  SYNC_RATIO_EPSILON,
  SYNC_STEP_PX,
  isSyncSignificant,
  ratioToScrollTop,
  scrollRatio,
} from '../src/renderer/preview/scroll-sync'

const view = (scrollTop: number, scrollHeight: number, clientHeight: number) => ({
  scrollTop,
  scrollHeight,
  clientHeight,
})

describe('scrollRatio', () => {
  it('maps the viewport position onto 0..1 proportionally', () => {
    // 1000px of scrollable range
    const s = view(250, 1500, 500)
    expect(scrollRatio(s)).toBeCloseTo(0.25)
    expect(scrollRatio(view(0, 1500, 500))).toBe(0)
    expect(scrollRatio(view(1000, 1500, 500))).toBeCloseTo(1)
  })

  it('is 0 when the pane cannot scroll (content shorter than the viewport)', () => {
    expect(scrollRatio(view(0, 400, 500))).toBe(0)
    expect(scrollRatio(view(0, 500, 500))).toBe(0)
    // a bogus scrollTop without scrollable range must not divide by zero
    expect(scrollRatio(view(120, 500, 500))).toBe(0)
  })

  it('clamps out-of-range positions into 0..1', () => {
    expect(scrollRatio(view(-40, 1500, 500))).toBe(0)
    expect(scrollRatio(view(9999, 1500, 500))).toBe(1)
  })
})

describe('ratioToScrollTop', () => {
  it("maps the proportion back onto the pane's own scroll range", () => {
    const s = view(0, 2000, 500)
    expect(ratioToScrollTop(0.25, s)).toBeCloseTo(375)
    expect(ratioToScrollTop(0, s)).toBe(0)
    expect(ratioToScrollTop(1, s)).toBe(1500)
  })

  it('stays at 0 for a pane with no scrollable range', () => {
    expect(ratioToScrollTop(0.75, view(0, 300, 500))).toBe(0)
  })

  it('clamps the incoming ratio', () => {
    const s = view(0, 2000, 500)
    expect(ratioToScrollTop(-3, s)).toBe(0)
    expect(ratioToScrollTop(2, s)).toBe(1500)
  })

  it('round-trips: applying the ratio of a position restores that position', () => {
    const s = view(333, 2000, 500)
    expect(ratioToScrollTop(scrollRatio(s), s)).toBeCloseTo(333)
  })
})

describe('isSyncSignificant', () => {
  it('rejects sub-pixel moves so the sync cannot hum in place', () => {
    expect(isSyncSignificant(500, 500.4)).toBe(false)
    expect(isSyncSignificant(500, 501, SYNC_STEP_PX)).toBe(true)
    // exact boundary counts as significant
    expect(isSyncSignificant(500, 501)).toBe(true)
  })
})

describe('ScrollSyncGate (loop protection)', () => {
  const now = 10_000

  it('silences the driven pane but never the other one', () => {
    const gate = new ScrollSyncGate()
    gate.drive('preview', now)
    // the programmatic reposition of the preview rings back as a scroll event
    expect(gate.shouldIgnore('preview', now + 10)).toBe(true)
    // a genuine source scroll one frame later must still go through
    expect(gate.shouldIgnore('source', now + 10)).toBe(false)
  })

  it('expires after the window, restoring two-way sync', () => {
    const gate = new ScrollSyncGate()
    gate.drive('source', now)
    expect(gate.shouldIgnore('source', now + SYNC_ECHO_WINDOW_MS - 1)).toBe(true)
    expect(gate.shouldIgnore('source', now + SYNC_ECHO_WINDOW_MS)).toBe(false)
  })

  it('refreshes the window on every drive of the same side', () => {
    const gate = new ScrollSyncGate()
    gate.drive('preview', now)
    gate.drive('preview', now + 100)
    expect(gate.shouldIgnore('preview', now + 100 + SYNC_ECHO_WINDOW_MS - 1)).toBe(true)
    expect(gate.shouldIgnore('preview', now + 100 + SYNC_ECHO_WINDOW_MS)).toBe(false)
  })

  it('lets the panes alternate: driving A does not block a later user scroll on B', () => {
    const gate = new ScrollSyncGate()
    gate.drive('preview', now)
    gate.drive('source', now + 50)
    expect(gate.shouldIgnore('source', now + 60)).toBe(true)
    expect(gate.shouldIgnore('preview', now + 60)).toBe(true)
    // after both windows expire everything passes again
    expect(gate.shouldIgnore('source', now + 300)).toBe(false)
    expect(gate.shouldIgnore('preview', now + 300)).toBe(false)
  })
})

describe('the ratio epsilon', () => {
  it('is far below any real user scroll but above scrollbar rounding', () => {
    // a full-height pane (1000px scrollable) moved by one wheel notch (~50px)
    expect(Math.abs(0.05) > SYNC_RATIO_EPSILON).toBe(true)
    // 1px on the same pane stays below it
    expect(1 / 1000 < SYNC_RATIO_EPSILON).toBe(true)
  })
})

import { describe, expect, it } from 'vitest'

import { createQuitFlow } from '../src/main/quit-flow'

/**
 * Quit bookkeeping (src/main/quit-flow.ts): how window closes decide to
 * persist the session during an app-wide quit vs an ordinary close, and how
 * a cancelled dirty-guard unwinds an in-flight quit (BUG-1104: the quitting
 * flag used to stay armed forever after a cancelled quit).
 */

describe('createQuitFlow', () => {
  it('starts in ordinary-close mode', () => {
    const flow = createQuitFlow()
    expect(flow.quitting).toBe(false)
    expect(flow.closeDecision(1)).toEqual({
      persist: true,
      skipStaged: true,
      excludeClosing: false,
    })
  })

  it('an ordinary close of one of several windows excludes the closing window', () => {
    const flow = createQuitFlow()
    expect(flow.closeDecision(2)).toEqual({
      persist: true,
      skipStaged: false,
      excludeClosing: true,
    })
  })

  it('a quit persists exactly once, at the first confirmed close, without skipStaged', () => {
    const flow = createQuitFlow()
    flow.begin()
    expect(flow.quitting).toBe(true)
    // first confirmed window close writes the quit snapshot — skipStaged stays
    // OFF (BUG-1105): windows whose dirty guards are still pending keep their
    // never-saved tabs restorable in case the quit is cancelled or the app
    // crashes mid-quit; the confirming window's staged files are already
    // deleted and simply prune away on the next restore
    expect(flow.closeDecision(3)).toEqual({
      persist: true,
      skipStaged: false,
      excludeClosing: false,
    })
    // ...later confirmed closes must not shrink it (their windows are gone)
    expect(flow.closeDecision(2)).toEqual({
      persist: false,
      skipStaged: false,
      excludeClosing: false,
    })
  })

  it('begin() re-arms a quit after the previous one was cancelled', () => {
    const flow = createQuitFlow()
    flow.begin()
    flow.closeDecision(1)
    flow.cancel()
    // a genuine second quit must start fresh: its first confirmed close
    // writes a snapshot again
    flow.begin()
    expect(flow.quitting).toBe(true)
    expect(flow.closeDecision(1).persist).toBe(true)
  })

  it('a repeated begin() mid-quit keeps persist-once armed (BUG-1219)', () => {
    const flow = createQuitFlow()
    flow.begin()
    expect(flow.closeDecision(3).persist).toBe(true)
    // a second before-quit while the quit is still in flight (repeat Cmd+Q
    // with a prompt open) must not re-arm the snapshot counter: the next
    // confirmed close used to overwrite the quit snapshot without the
    // already-closed windows
    flow.begin()
    expect(flow.quitting).toBe(true)
    expect(flow.closeDecision(2).persist).toBe(false)
  })

  it('a cancel after a repeated begin() still requests the recovery rewrite (BUG-1219)', () => {
    const flow = createQuitFlow()
    flow.begin()
    flow.closeDecision(2)
    flow.begin()
    // the snapshot from before the repeat is stale (a window already
    // closed); the cancel must replace it, not skip the rewrite like it did
    // when the repeat had reset the counter
    expect(flow.cancel()).toBe(true)
    expect(flow.quitting).toBe(false)
  })

  it('a cancelled guard unwinds the quit: later closes are ordinary again (BUG-1104)', () => {
    const flow = createQuitFlow()
    flow.begin()
    // nothing persisted yet — the cancel needs no session restore
    expect(flow.cancel()).toBe(false)
    expect(flow.quitting).toBe(false)
    // the very next ordinary close must take the per-window branch, not the
    // quit branch (the original bug: these stayed quit-closes forever)
    expect(flow.closeDecision(2)).toEqual({
      persist: true,
      skipStaged: false,
      excludeClosing: true,
    })
  })

  it('a cancel after the quit snapshot landed asks for a live re-serialize', () => {
    const flow = createQuitFlow()
    flow.begin()
    flow.closeDecision(2)
    // a second window cancelled its guard after the first one confirmed:
    // the quit-time write must be replaced from the surviving windows
    expect(flow.cancel()).toBe(true)
    expect(flow.quitting).toBe(false)
    expect(flow.closeDecision(1)).toEqual({
      persist: true,
      skipStaged: true,
      excludeClosing: false,
    })
  })

  it('cancel() on a flow that never quit is a no-op', () => {
    const flow = createQuitFlow()
    expect(flow.cancel()).toBe(false)
    expect(flow.quitting).toBe(false)
  })

  it('a failed snapshot write re-arms persist-once so the next close retries (BUG-1224)', () => {
    const flow = createQuitFlow()
    flow.begin()
    expect(flow.closeDecision(3).persist).toBe(true)
    // the write that closeDecision armed just failed (disk full, read-only
    // userData): without the re-arm every later confirmed close would skip
    // its retry and the next launch would resurrect already-closed windows
    // from the stale snapshot
    flow.markSnapshotWriteFailed()
    expect(flow.closeDecision(2).persist).toBe(true)
  })

  it('a cancel right after a failed snapshot write needs no recovery rewrite (BUG-1224)', () => {
    const flow = createQuitFlow()
    flow.begin()
    flow.closeDecision(3)
    flow.markSnapshotWriteFailed()
    // nothing actually landed, so the cancel skips the (equally doomed)
    // recovery rewrite
    expect(flow.cancel()).toBe(false)
    expect(flow.quitting).toBe(false)
  })

  it('markSnapshotWriteFailed is a no-op outside a quit', () => {
    const flow = createQuitFlow()
    flow.markSnapshotWriteFailed()
    expect(flow.quitting).toBe(false)
    // an ordinary close is unaffected (its write failures retry via the
    // debounced save; there is no persist-once to re-arm)
    expect(flow.closeDecision(1).persist).toBe(true)
  })
})

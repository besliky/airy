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
    expect(flow.closeDecision(1)).toEqual({ persist: true, skipStaged: true, excludeClosing: false })
  })

  it('an ordinary close of one of several windows excludes the closing window', () => {
    const flow = createQuitFlow()
    expect(flow.closeDecision(2)).toEqual({
      persist: true,
      skipStaged: false,
      excludeClosing: true,
    })
  })

  it('a quit persists exactly once, at the first confirmed close', () => {
    const flow = createQuitFlow()
    flow.begin()
    expect(flow.quitting).toBe(true)
    // first confirmed window close writes the quit snapshot...
    expect(flow.closeDecision(3)).toEqual({ persist: true, skipStaged: true, excludeClosing: false })
    // ...later confirmed closes must not shrink it (their windows are gone)
    expect(flow.closeDecision(2)).toEqual({ persist: false, skipStaged: false, excludeClosing: false })
  })

  it('begin() re-arms a quit after the previous one finished', () => {
    const flow = createQuitFlow()
    flow.begin()
    flow.closeDecision(1)
    flow.begin()
    expect(flow.closeDecision(1).persist).toBe(true)
  })

  it('a cancelled guard unwinds the quit: later closes are ordinary again (BUG-1104)', () => {
    const flow = createQuitFlow()
    flow.begin()
    // nothing persisted yet — the cancel needs no session restore
    expect(flow.cancel()).toBe(false)
    expect(flow.quitting).toBe(false)
    // the very next ordinary close must take the per-window branch, not the
    // quit branch (the original bug: these stayed quit-closes forever)
    expect(flow.closeDecision(2)).toEqual({ persist: true, skipStaged: false, excludeClosing: true })
  })

  it('a cancel after the quit snapshot landed asks for a live re-serialize', () => {
    const flow = createQuitFlow()
    flow.begin()
    flow.closeDecision(2)
    // a second window cancelled its guard after the first one confirmed:
    // the quit-time write must be replaced from the surviving windows
    expect(flow.cancel()).toBe(true)
    expect(flow.quitting).toBe(false)
    expect(flow.closeDecision(1)).toEqual({ persist: true, skipStaged: true, excludeClosing: false })
  })

  it('cancel() on a flow that never quit is a no-op', () => {
    const flow = createQuitFlow()
    expect(flow.cancel()).toBe(false)
    expect(flow.quitting).toBe(false)
  })
})

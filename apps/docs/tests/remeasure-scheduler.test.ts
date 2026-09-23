/**
 * The remeasure scheduler drives the O(document) pagination pass. Its budgets
 * must be deterministic (fake timers, injectable clock) and its queue must
 * converge: however often callers request, the pass runs a bounded number of
 * times (PERF-1639 — the pre-fix fixed 300ms debounce re-ran the pass over a
 * growing document and never settled).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  REMEASURE_MAX_DELAY_MS,
  REMEASURE_MIN_DELAY_MS,
  createRemeasureScheduler,
} from '../src/renderer/remeasure-scheduler'

describe('createRemeasureScheduler', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })
  afterEach(() => {
    vi.useRealTimers()
  })

  it('coalesces any number of requests into one pending run', () => {
    const run = vi.fn()
    const scheduler = createRemeasureScheduler({ run })
    for (let i = 0; i < 50; i++) scheduler.request()
    expect(run).not.toHaveBeenCalled()
    expect(scheduler.pending()).toBe(true)
    vi.advanceTimersByTime(REMEASURE_MIN_DELAY_MS)
    expect(run).toHaveBeenCalledTimes(1)
    expect(scheduler.pending()).toBe(false)
  })

  it('a request during the delay defers the run to the latest request (trailing debounce)', () => {
    const run = vi.fn()
    const scheduler = createRemeasureScheduler({ run })
    scheduler.request()
    vi.advanceTimersByTime(200)
    scheduler.request() // resets the tail
    vi.advanceTimersByTime(200)
    expect(run).not.toHaveBeenCalled()
    vi.advanceTimersByTime(100)
    expect(run).toHaveBeenCalledTimes(1)
  })

  it('cheap runs keep the minimal delay', () => {
    const run = vi.fn()
    const scheduler = createRemeasureScheduler({ run, now: () => 0 })
    scheduler.request()
    vi.advanceTimersByTime(REMEASURE_MIN_DELAY_MS)
    expect(scheduler.nextDelayMs()).toBe(REMEASURE_MIN_DELAY_MS)
  })

  it('expensive runs back off proportionally, clamped to the max delay', () => {
    let clock = 0
    const run = vi.fn(() => {
      clock += 1000 // each pass costs 1s of work
    })
    const scheduler = createRemeasureScheduler({ run, now: () => clock })
    scheduler.request()
    vi.advanceTimersByTime(REMEASURE_MIN_DELAY_MS)
    // 4x the measured cost, clamped to the maximum
    expect(scheduler.lastRunMs()).toBe(1000)
    expect(scheduler.nextDelayMs()).toBe(REMEASURE_MAX_DELAY_MS)
    // a moderate cost stays inside the 4x rule
    run.mockImplementation(() => {
      clock += 500
    })
    scheduler.request()
    vi.advanceTimersByTime(REMEASURE_MAX_DELAY_MS)
    expect(scheduler.nextDelayMs()).toBe(2000)
  })

  it('converges: an edit storm runs the pass a bounded number of times', () => {
    const run = vi.fn()
    const scheduler = createRemeasureScheduler({ run, now: () => 0 })
    // 500 keystrokes at 50ms intervals: the pass runs only in the quiet tails
    for (let i = 0; i < 500; i++) {
      scheduler.request()
      vi.advanceTimersByTime(50)
    }
    const duringStorm = run.mock.calls.length
    expect(duringStorm).toBeLessThan(500)
    // after the storm stops, exactly one final run settles the queue
    vi.advanceTimersByTime(REMEASURE_MAX_DELAY_MS * 2)
    expect(run.mock.calls.length).toBe(duringStorm + 1)
    expect(scheduler.pending()).toBe(false)
  })

  it('cancel drops the pending run', () => {
    const run = vi.fn()
    const scheduler = createRemeasureScheduler({ run })
    scheduler.request()
    scheduler.cancel()
    vi.advanceTimersByTime(REMEASURE_MAX_DELAY_MS * 2)
    expect(run).not.toHaveBeenCalled()
    expect(scheduler.pending()).toBe(false)
  })
})

import { describe, expect, it, vi } from 'vitest'
import {
  createQueuedWorkbookDelivery,
  type DeliveryScheduler,
} from '../src/main/queued-workbook-delivery'

/**
 * Queued-workbook delivery state machine (pure, injected timers): the initial
 * send, the bounded resends when no menu-ready signal arrives, the ready
 * flush, and the stop conditions (workbook consumed / tab gone).
 */
function manualScheduler() {
  const pending: { fn: () => void; ms: number; cancelled: boolean }[] = []
  const schedule: DeliveryScheduler = (fn, ms) => {
    const entry = { fn, ms, cancelled: false }
    pending.push(entry)
    return () => {
      entry.cancelled = true
    }
  }
  const fireNext = () => {
    const entry = pending.find((candidate) => !candidate.cancelled)
    if (!entry) throw new Error('no pending timer')
    entry.cancelled = true
    entry.fn()
  }
  const pendingCount = () => pending.filter((candidate) => !candidate.cancelled).length
  return { schedule, fireNext, pendingCount }
}

describe('createQueuedWorkbookDelivery', () => {
  it('sends once immediately and resends at most maxResends times without a ready signal', () => {
    const { schedule, fireNext, pendingCount } = manualScheduler()
    const sendOpen = vi.fn()
    const delivery = createQueuedWorkbookDelivery({
      sendOpen,
      isStillWaiting: () => true,
      maxResends: 2,
      delayMs: 2_000,
      schedule,
    })

    delivery.start()
    expect(sendOpen).toHaveBeenCalledTimes(1)
    expect(pendingCount()).toBe(1)

    fireNext()
    expect(sendOpen).toHaveBeenCalledTimes(2)
    fireNext()
    expect(sendOpen).toHaveBeenCalledTimes(3)

    // both resends spent: no timer is pending anymore
    expect(pendingCount()).toBe(0)
    expect(() => fireNext()).toThrow('no pending timer')
  })

  it('stops resending once the workbook was consumed or the tab is gone', () => {
    const { schedule, fireNext } = manualScheduler()
    const sendOpen = vi.fn()
    let waiting = true
    const delivery = createQueuedWorkbookDelivery({
      sendOpen,
      isStillWaiting: () => waiting,
      maxResends: 2,
      delayMs: 2_000,
      schedule,
    })

    delivery.start()
    waiting = false // e.g. the renderer consumed the queued path
    fireNext()
    expect(sendOpen).toHaveBeenCalledTimes(1)
  })

  it('flushes once on the ready signal and stops resending', () => {
    const { schedule, fireNext } = manualScheduler()
    const sendOpen = vi.fn()
    const delivery = createQueuedWorkbookDelivery({
      sendOpen,
      isStillWaiting: () => true,
      maxResends: 2,
      delayMs: 2_000,
      schedule,
    })

    delivery.start()
    delivery.onReady()
    expect(sendOpen).toHaveBeenCalledTimes(2)

    // the resend timer was cancelled by onReady
    expect(() => fireNext()).toThrow('no pending timer')
  })

  it('does not flush on ready when nothing is waiting anymore', () => {
    const sendOpen = vi.fn()
    const delivery = createQueuedWorkbookDelivery({
      sendOpen,
      isStillWaiting: () => false,
      maxResends: 2,
      delayMs: 2_000,
      schedule: () => () => {},
    })

    delivery.start()
    delivery.onReady()
    expect(sendOpen).toHaveBeenCalledTimes(1)
  })

  it('a fresh start cancels the previous resend chain', () => {
    const { schedule, fireNext, pendingCount } = manualScheduler()
    const sendOpen = vi.fn()
    const delivery = createQueuedWorkbookDelivery({
      sendOpen,
      isStillWaiting: () => true,
      maxResends: 2,
      delayMs: 2_000,
      schedule,
    })

    delivery.start()
    expect(pendingCount()).toBe(1)
    delivery.start()
    expect(pendingCount()).toBe(1) // old timer replaced, not stacked
    fireNext()
    // two immediate sends (one per start) + one resend fired from the new chain
    expect(sendOpen).toHaveBeenCalledTimes(3)
  })
})

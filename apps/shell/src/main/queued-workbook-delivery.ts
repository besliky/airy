/**
 * Delivers the shell-queued workbook's 'open' menu action to the sheets
 * renderer. The renderer subscribes to menu actions only after Univer mounts
 * (seconds on cold start), so the first send may land before the subscription
 * exists: the renderer's one-time menu-ready signal flushes it immediately,
 * and a bounded resend loop covers a stale preload that never sends the
 * signal. Extracted from index.ts with injected timers so the state machine
 * is unit-testable.
 */
export interface QueuedWorkbookDelivery {
  /** (Re)start delivery: send once now, then arm the bounded resends. */
  start(): void
  /** Renderer menu-ready signal: flush once if still waiting, stop resending. */
  onReady(): void
  /** Stop resending (workbook consumed, tab gone, app quitting). */
  cancel(): void
}

export interface DeliveryScheduler {
  /** Run `fn` after `ms`; the returned function cancels it. */
  (fn: () => void, ms: number): () => void
}

const defaultSchedule: DeliveryScheduler = (fn, ms) => {
  const timer = setTimeout(fn, ms)
  return () => clearTimeout(timer)
}

export function createQueuedWorkbookDelivery(deps: {
  sendOpen: () => void
  /** Is the queued workbook still unconsumed and its tab alive? */
  isStillWaiting: () => boolean
  /** Resend attempts after the initial send when no ready signal arrives. */
  maxResends?: number
  delayMs?: number
  schedule?: DeliveryScheduler
}): QueuedWorkbookDelivery {
  const maxResends = deps.maxResends ?? 2
  const delayMs = deps.delayMs ?? 2_000
  const schedule = deps.schedule ?? defaultSchedule
  let cancelPending: (() => void) | null = null
  let resendsLeft = 0

  function cancel(): void {
    cancelPending?.()
    cancelPending = null
  }

  function arm(): void {
    if (resendsLeft <= 0) return
    cancelPending = schedule(() => {
      cancelPending = null
      if (!deps.isStillWaiting()) return
      deps.sendOpen()
      resendsLeft -= 1
      arm()
    }, delayMs)
  }

  return {
    start() {
      cancel()
      deps.sendOpen()
      resendsLeft = maxResends
      arm()
    },
    onReady() {
      if (deps.isStillWaiting()) deps.sendOpen()
      cancel()
    },
    cancel,
  }
}

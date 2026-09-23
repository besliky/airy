/**
 * Self-tuning scheduler for the full-document remeasure pass (pagination,
 * page gaps, margin annotations). The pass costs O(document), so on large
 * documents it must not run on every keystroke's 300ms tail: the delay after
 * each run adapts to that run's measured cost (PERF-1639). Keystrokes still
 * land immediately — only the expensive layout pass is deferred.
 *
 * Deterministic by construction: inject setTimeout/clearTimeout/now in tests
 * (fake timers) and assert both the budget bounds and that repeated requests
 * converge to a bounded number of runs.
 */

export interface RemeasureSchedulerOptions {
  /** the expensive pass */
  run: () => void
  /** shortest delay between request and run (fast-document pacing) */
  minDelayMs?: number
  /** longest a run may be deferred while input keeps coming */
  maxDelayMs?: number
  /** a run is considered expensive above this duration; delay backs off toward maxDelayMs */
  expensiveMs?: number
  setTimeout?: (cb: () => void, ms: number) => unknown
  clearTimeout?: (id: unknown) => void
  now?: () => number
}

export const REMEASURE_MIN_DELAY_MS = 300
export const REMEASURE_MAX_DELAY_MS = 3000
export const REMEASURE_EXPENSIVE_MS = 200

/**
 * Multiple of the measured run duration used as the next delay, so a pass that
 * costs seconds is re-run at most every few seconds while edits keep arriving.
 */
const DELAY_PER_RUN_MS = 4

export interface RemeasureScheduler {
  /** request a run; trailing-debounce semantics with the adaptive delay */
  request: () => void
  /** drop a pending run (the pass itself always stays schedulable) */
  cancel: () => void
  /** true while a run is scheduled */
  pending: () => boolean
  /** delay the next request would use */
  nextDelayMs: () => number
  /** duration of the last executed run (0 before the first run) */
  lastRunMs: () => number
}

export function createRemeasureScheduler(options: RemeasureSchedulerOptions): RemeasureScheduler {
  const {
    run,
    minDelayMs = REMEASURE_MIN_DELAY_MS,
    maxDelayMs = REMEASURE_MAX_DELAY_MS,
    expensiveMs = REMEASURE_EXPENSIVE_MS,
    setTimeout: setT = (cb, ms) => window.setTimeout(cb, ms),
    clearTimeout: clearT = (id) => window.clearTimeout(id as number),
    now = () => performance.now(),
  } = options

  let timer: unknown = null
  let last = 0

  const nextDelay = (): number => {
    if (last < expensiveMs) return minDelayMs
    return Math.min(maxDelayMs, Math.max(minDelayMs, Math.round(last * DELAY_PER_RUN_MS)))
  }

  const scheduled = () => {
    timer = null
    const t0 = now()
    try {
      run()
    } finally {
      last = Math.max(0, now() - t0)
    }
  }

  return {
    request: () => {
      if (timer !== null) clearT(timer)
      timer = setT(scheduled, nextDelay())
    },
    cancel: () => {
      if (timer !== null) {
        clearT(timer)
        timer = null
      }
    },
    pending: () => timer !== null,
    nextDelayMs: nextDelay,
    lastRunMs: () => last,
  }
}

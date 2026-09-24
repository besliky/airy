import type { RecentEntry } from '../../shared/home-api'
import { HOME_PATHS_CAP } from '../../shared/home-paths'

/**
 * Project-catalog loading for the Home screen (BUG-1676): a project's full
 * path list is stat-ed in bounded chunks with a main-thread yield between
 * chunks, so even a 20k-file catalog never blocks the renderer in one long
 * task and rows appear progressively while the load runs. Pure orchestration
 * (every effectful step is injected), so chunking, ordering, the yield
 * cadence and the clamp at HOME_PATHS_CAP are all unit-testable.
 */

/** paths per statPaths round-trip; small enough to keep each IPC call and each
 * state update cheap, large enough that a 20k catalog needs only 40 calls */
export const STAT_CHUNK_SIZE = 500

export type ProjectCountKey = 'fileCount' | 'fileCountOne' | 'fileCountOver'

/**
 * i18n key for the project file counter. Beyond HOME_PATHS_CAP the loaded
 * list is a clamp of the real catalog, so the count is shown as "{n}+"
 * (fileCountOver) instead of silently under-reporting (BUG-1676).
 */
export function projectCountKey(count: number, overCap: boolean): ProjectCountKey {
  if (overCap) return 'fileCountOver'
  return count === 1 ? 'fileCountOne' : 'fileCount'
}

/** newest-first base order for the project table (the header sort flips it). */
export function sortByModifiedDesc(entries: readonly RecentEntry[]): RecentEntry[] {
  return [...entries].sort((a, b) => b.mtimeMs - a.mtimeMs)
}

/**
 * Hand control back to the event loop between chunks so pending input,
 * paint and other tasks interleave with the load. setImmediate posts a fresh
 * macrotask (Node / Electron renderer); a 0ms timer is the browser fallback.
 */
export function yieldToMain(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof setImmediate === 'function') setImmediate(resolve)
    else setTimeout(resolve, 0)
  })
}

export interface StatPathsChunkedDeps {
  /** per-chunk stat round-trip (window.aiOffice.statPaths in the app) */
  statPaths: (paths: string[]) => Promise<RecentEntry[]>
  /** test seam; defaults to yieldToMain */
  yield?: () => Promise<void>
}

/**
 * Stat a whole path list chunk by chunk. `onProgress` fires after every
 * chunk with the entries accumulated so far (newest-first), letting the
 * caller render progressively; the resolved value is the full clamped list.
 * The input is clamped to HOME_PATHS_CAP — the same bound the main-process
 * handler enforces — so callers can treat `paths.length > cap` as the
 * over-cap signal for the counter.
 */
export async function statPathsChunked(
  paths: readonly string[],
  deps: StatPathsChunkedDeps,
  onProgress: (entries: RecentEntry[]) => void,
): Promise<RecentEntry[]> {
  const bounded = paths.length > HOME_PATHS_CAP ? paths.slice(0, HOME_PATHS_CAP) : paths
  const wait = deps.yield ?? yieldToMain
  const collected: RecentEntry[] = []
  for (let start = 0; start < bounded.length; start += STAT_CHUNK_SIZE) {
    if (start > 0) await wait()
    const chunk = bounded.slice(start, start + STAT_CHUNK_SIZE)
    const stats = await deps.statPaths(chunk)
    for (const entry of stats) collected.push(entry)
    onProgress(sortByModifiedDesc(collected))
  }
  // same newest-first order the progress callbacks deliver — one contract
  return sortByModifiedDesc(collected)
}

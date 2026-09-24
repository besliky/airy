/**
 * Bounded path-list parsing for the home:stat-paths channel. It stats every
 * entry it is handed, so an unbounded list would mean unbounded filesystem
 * work per call. The cap covers recent.json's 100 entries, starred files and
 * drag selections with headroom, AND full project catalogs: a project's file
 * list is loaded through this channel, so the cap is the largest project the
 * Home screen can index (beyond it the renderer shows an honest "{n}+ files"
 * counter instead of a silently truncated list — BUG-1676). Callers that load
 * a whole catalog chunk their requests (renderer project-files loader), so
 * the per-call cap only guards against absurd IPC payloads. Shared by main
 * (the handler's guard) and the renderer (the loader + counter), so both
 * sides agree on one number. Pure so the cap is unit-tested.
 */
export const HOME_PATHS_CAP = 20_000

/** strings only, extras beyond the cap ignored (first `cap` win). */
export function stringPathsCapped(value: unknown, cap: number = HOME_PATHS_CAP): string[] {
  if (!Array.isArray(value) || cap < 1) return []
  const paths: string[] = []
  for (const entry of value) {
    if (typeof entry !== 'string') continue
    if (paths.length >= cap) break
    paths.push(entry)
  }
  return paths
}

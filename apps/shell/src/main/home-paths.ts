/**
 * Bounded path-list parsing for the home:stat-paths channel. It stats every
 * entry it is handed, so an unbounded list would mean unbounded filesystem
 * work per call; the cap (256) covers recent.json's 100 entries with
 * headroom for starred files and drag selections. Pure so the cap is
 * unit-tested.
 */
export const HOME_PATHS_CAP = 256

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

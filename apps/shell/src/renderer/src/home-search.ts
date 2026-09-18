import type { RecentEntry } from '../../shared/home-api'

/**
 * Live search support for the Home file lists: folding, the name filter the
 * visible tables use, and the corpus merge across the Recent / Starred /
 * project-files sources. Pure string/array logic only — no React, no IPC —
 * so every rule is unit-testable.
 */

/**
 * Fold text for case- and accent-insensitive comparison: NFD-decompose, drop
 * combining marks, lowercase. Built-ins only (no fold dependency), and CJK
 * text has no decomposition, so Chinese/Japanese/Korean names match verbatim.
 */
export function foldForSearch(value: string): string {
  return value.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase()
}

/** whether the query engages the filter at all (anything but whitespace) */
export function isSearchActive(query: string): boolean {
  return query.trim() !== ''
}

/**
 * Live file-name filter for the home tables. A blank query means "no
 * filter": the input array is returned untouched (same reference) so the
 * plain view keeps its existing behavior. Otherwise entries whose folded
 * NAME contains the folded query survive; the input order is preserved.
 * The path is deliberately not searched — the field is a name filter.
 */
export function filterFileEntries<T extends { name: string }>(
  entries: readonly T[],
  query: string,
): T[] {
  const needle = foldForSearch(query.trim())
  if (needle === '') return entries as T[]
  return entries.filter((entry) => foldForSearch(entry.name).includes(needle))
}

/**
 * Combine the home lists (recent, starred, open-project files) into one
 * searchable corpus, deduped by path: an earlier list wins, and every list
 * keeps its internal order — the same first-seen ordering the visible
 * tables already use when two lists carry the same file.
 */
export function mergeFileLists(...lists: ReadonlyArray<readonly RecentEntry[]>): RecentEntry[] {
  const seen = new Set<string>()
  const merged: RecentEntry[] = []
  for (const list of lists) {
    for (const entry of list) {
      if (seen.has(entry.path)) continue
      seen.add(entry.path)
      merged.push(entry)
    }
  }
  return merged
}

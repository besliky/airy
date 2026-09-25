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
 * Pre-folded name index for live search (PERF-1736): the previous per-keystroke
 * path re-folded (NFD + mark-strip + lowercase) every visible name on every
 * keystroke. The index folds each name once per corpus build and lets each
 * keystroke run a plain `includes` sweep, keeping the filter itself at a few
 * milliseconds even at the 20k-file catalog cap.
 */
export interface NameFilterIndex<T> {
  /** the indexed corpus, kept so a blank query can return it untouched */
  entries: readonly T[]
  /** folded name per entry, same order as `entries` */
  foldedNames: string[]
}

/** fold every name once; call when the corpus array identity changes */
export function buildNameFilterIndex<T extends { name: string }>(
  entries: readonly T[],
): NameFilterIndex<T> {
  return { entries, foldedNames: entries.map((entry) => foldForSearch(entry.name)) }
}

/**
 * Live filter over a pre-built index, with the exact semantics of
 * `filterFileEntries`: blank query returns the corpus untouched (same
 * reference), otherwise entries whose folded name contains the folded query
 * survive, input order preserved, path never searched.
 */
export function filterWithIndex<T>(index: NameFilterIndex<T>, query: string): T[] {
  const needle = foldForSearch(query.trim())
  if (needle === '') return index.entries as T[]
  const { entries, foldedNames } = index
  const hits: T[] = []
  for (let i = 0; i < entries.length; i++) {
    if (foldedNames[i]?.includes(needle)) hits.push(entries[i] as T)
  }
  return hits
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

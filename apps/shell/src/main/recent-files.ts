import { stat } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import type { RecentEntry, RecentPage, RecentQuery } from '../shared/home-api'

const RECENT_PAGE_DEFAULT = 50
const RECENT_PAGE_MAX = 200

/**
 * Short-TTL stat cache: the Home screen re-queries recents on every window
 * focus, and stat-ing up to 100 paths (network drives are the worst case)
 * made each focus event expensive. Entries go stale within TTL_MS, so a file
 * changed right after a query can show a few seconds old mtime/size until the
 * next refresh — the list itself always reflects the store immediately.
 */
const STAT_CACHE_TTL_MS = 5_000
interface CachedStat {
  mtimeMs: number
  sizeBytes: number
  missing: boolean
  at: number
}
const statCache = new Map<string, CachedStat>()

/** Test hook: drop cached stats so a fresh query re-stats everything. */
export function clearRecentStatCache(): void {
  statCache.clear()
}

function toRecentEntry(
  path: string,
  starredPaths: ReadonlySet<string>,
  cached: CachedStat,
): RecentEntry {
  // A failed stat is often transient (disconnected drive, pending mount,
  // cloud placeholder) — dropping the entry made the recents list silently
  // lose files until a later reload (r158). Word keeps unavailable recents
  // listed; the row is flagged so the UI can dim it and offer removal.
  if (cached.missing) {
    return {
      path,
      name: basename(path),
      ext: extname(path).slice(1).toLowerCase(),
      mtimeMs: 0,
      sizeBytes: 0,
      starred: starredPaths.has(path),
      missing: true,
    }
  }
  return {
    path,
    name: basename(path),
    ext: extname(path).slice(1).toLowerCase(),
    mtimeMs: cached.mtimeMs,
    sizeBytes: cached.sizeBytes,
    starred: starredPaths.has(path),
  }
}

async function statCached(path: string): Promise<CachedStat> {
  const cached = statCache.get(path)
  if (cached && Date.now() - cached.at < STAT_CACHE_TTL_MS) return cached
  const fresh: CachedStat = await stat(path)
    .then((info) => ({
      mtimeMs: info.mtimeMs,
      sizeBytes: info.size,
      missing: false,
      at: Date.now(),
    }))
    .catch(
      () =>
        ({
          mtimeMs: 0,
          sizeBytes: 0,
          missing: true,
          at: Date.now(),
        }) satisfies CachedStat,
    )
  statCache.set(path, fresh)
  return fresh
}

export async function statPathEntries(
  paths: readonly string[],
  starredPaths: ReadonlySet<string>,
): Promise<RecentEntry[]> {
  const stats = await Promise.all(paths.map((path) => statCached(path)))
  return paths.map((path, index) => toRecentEntry(path, starredPaths, stats[index]!))
}

export function normalizeRecentQuery(
  raw: unknown,
): Required<Omit<RecentQuery, 'ext'>> & { ext?: string } {
  const query = (raw ?? {}) as RecentQuery
  const offset = Number.isFinite(query.offset) ? Math.max(0, Math.floor(query.offset!)) : 0
  const limit = Number.isFinite(query.limit)
    ? Math.min(RECENT_PAGE_MAX, Math.max(0, Math.floor(query.limit!)))
    : RECENT_PAGE_DEFAULT
  // Sidebar keys are bare extensions ("xlsx"), but IPC callers may send
  // ".xlsx", " XLSX ", or "..." — normalize so openable files cannot hide
  // behind a filter that only differs in dots/case/whitespace.
  const rawExt =
    typeof query.ext === 'string' ? query.ext.trim().toLowerCase().replace(/^\.+/, '') : ''
  const ext = rawExt ? rawExt : undefined
  return { offset, limit, ext }
}

/** sidebar filter keys that stand for a family of extensions, not one exact ext */
const EXT_FAMILY: Record<string, readonly string[]> = {
  xlsx: ['xlsx', 'xlsm', 'xls'],
  html: ['html', 'htm'],
}

/** Page over the recents paths, preserving the source's newest-first order (unavailable paths stay, flagged missing). */
export async function pageRecentPaths(
  paths: readonly string[],
  raw: unknown,
  starredPaths: ReadonlySet<string>,
): Promise<RecentPage> {
  const { offset, limit, ext } = normalizeRecentQuery(raw)
  const all = await statPathEntries(paths, starredPaths)
  const family = ext ? (EXT_FAMILY[ext] ?? [ext]) : undefined
  const filtered = family ? all.filter((entry) => family.includes(entry.ext)) : all
  return {
    entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
    total: filtered.length,
    totalAll: all.length,
  }
}

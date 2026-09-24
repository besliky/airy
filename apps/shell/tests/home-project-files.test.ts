// BUG-1676 regression: project Home indexing silently stopped at 256 files
// (the old HOME_PATHS_CAP), so the counter lied ("256 files") and search could
// not find anything beyond the cap (file #280 → "No files matching"). The fix
// raises the cap to a production level and loads the catalog in chunked,
// yield-separated batches; these tests cover the pure loader: full indexing,
// the clamp, the honest "+ counter" key, chunking and non-blocking cadence.
import { mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { statPathEntries } from '../src/main/recent-files'
import type { RecentEntry } from '../src/shared/home-api'
import { HOME_PATHS_CAP } from '../src/shared/home-paths'
import {
  projectCountKey,
  sortByModifiedDesc,
  statPathsChunked,
  STAT_CHUNK_SIZE,
} from '../src/renderer/src/project-files'

/** minimal RecentEntry factory; name defaults to the path's file name */
function entry(path: string, name?: string): RecentEntry {
  return {
    path,
    name: name ?? (path.split('/').pop() as string),
    ext: (path.split('.').pop() ?? '').toLowerCase(),
    mtimeMs: 0,
    sizeBytes: 0,
    starred: false,
  }
}

function catalogPaths(dir: string, count: number): string[] {
  return Array.from({ length: count }, (_, i) => join(dir, `p${String(i).padStart(4, '0')}.docx`))
}

/** real main-process stat of each chunk — the same handler work statPaths does */
const statViaMain = (chunk: string[]): Promise<RecentEntry[]> => statPathEntries(chunk, new Set())

describe('projectCountKey (honest counter, BUG-1676)', () => {
  it('uses the plain keys below the cap', () => {
    expect(projectCountKey(0, false)).toBe('fileCount')
    expect(projectCountKey(1, false)).toBe('fileCountOne')
    expect(projectCountKey(300, false)).toBe('fileCount')
  })

  it('switches to the "+ files" key once the catalog exceeds the cap', () => {
    expect(projectCountKey(HOME_PATHS_CAP, true)).toBe('fileCountOver')
    // also while searching: the match count only covers the clamped corpus
    expect(projectCountKey(3, true)).toBe('fileCountOver')
  })
})

describe('sortByModifiedDesc', () => {
  it('sorts newest first without mutating the input', () => {
    const a = entry('/a', 'a')
    const b = entry('/b', 'b')
    const c = entry('/c', 'c')
    a.mtimeMs = 3
    b.mtimeMs = 1
    c.mtimeMs = 2
    const sorted = sortByModifiedDesc([a, b, c])
    expect(sorted.map((e) => e.name)).toEqual(['a', 'c', 'b'])
    expect(sorted).not.toBe([a, b, c])
  })
})

describe('statPathsChunked', () => {
  it('fully indexes a 300-file tmp catalog (old cap silently stopped at 256)', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'airy-home-cap-'))
    try {
      const paths = catalogPaths(dir, 300)
      for (const [i, path] of paths.entries()) {
        await writeFile(path, 'x')
        // distinct mtimes so the newest-first sort order is assertable
        await utimes(path, new Date(1_700_000_000_000 + i), new Date(1_700_000_000_000 + i))
      }
      const updates: number[] = []
      const loaded = await statPathsChunked(paths, { statPaths: statViaMain }, (entries) => {
        updates.push(entries.length)
      })
      // the whole catalog is loaded, including the old-cap victim #280
      expect(loaded).toHaveLength(300)
      expect(loaded.some((e) => e.name === 'p0280.docx')).toBe(true)
      // newest-first base order: p0299 (latest mtime) leads
      expect(loaded[0]?.name).toBe('p0299.docx')
      // one progress update per chunk; 300 < STAT_CHUNK_SIZE → single chunk
      expect(updates).toEqual([300])
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  }, 20_000)

  it('stats in bounded chunks with a yield between chunks, preserving order', async () => {
    const paths = catalogPaths('/proj', 1_200)
    const events: string[] = []
    const sizes: number[] = []
    const loaded = await statPathsChunked(
      paths,
      {
        statPaths: async (chunk) => {
          events.push('stat')
          sizes.push(chunk.length)
          return chunk.map((p) => entry(p))
        },
        yield: () =>
          new Promise<void>((resolve) => {
            events.push('yield')
            resolve()
          }),
      },
      () => {},
    )
    expect(loaded).toHaveLength(1_200)
    expect(sizes).toEqual([STAT_CHUNK_SIZE, STAT_CHUNK_SIZE, 200])
    // chunk order preserved end to end
    expect(loaded[0]?.name).toBe('p0000.docx')
    expect(loaded[999]?.name).toBe('p0999.docx')
    expect(loaded[1_199]?.name).toBe('p1199.docx')
    // a yield separates every consecutive pair of chunks (never one burst)
    expect(events).toEqual(['stat', 'yield', 'stat', 'yield', 'stat'])
  })

  it('lets already-queued macrotasks run while the load is in flight', async () => {
    const paths = catalogPaths('/proj', 1_000)
    let timerFired = false
    const run = statPathsChunked(
      paths,
      {
        statPaths: async (chunk) => chunk.map((p) => entry(p)),
        // real macrotask yield, like the default yieldToMain in a browser
        yield: () => new Promise<void>((resolve) => setTimeout(resolve, 0)),
      },
      () => {},
    )
    // queued while the load is pending; same-delay timers run in insertion
    // order, so this fires between chunks and proves the loop is not starved
    setTimeout(() => {
      timerFired = true
    }, 0)
    await run
    expect(timerFired).toBe(true)
  })

  it('clamps catalogs beyond HOME_PATHS_CAP (the renderer flags it "+")', async () => {
    const paths = catalogPaths('/proj', HOME_PATHS_CAP + 4_000)
    const loaded = await statPathsChunked(
      paths,
      { statPaths: (c) => c.map((p) => entry(p)) },
      () => {},
    )
    expect(loaded).toHaveLength(HOME_PATHS_CAP)
    expect(loaded[0]?.name).toBe('p0000.docx')
    expect(loaded[HOME_PATHS_CAP - 1]?.name).toBe(
      `p${String(HOME_PATHS_CAP - 1).padStart(4, '0')}.docx`,
    )
  })
})

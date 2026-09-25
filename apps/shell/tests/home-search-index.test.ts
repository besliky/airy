// PERF-1736: the pre-folded name index. Correctness must match the plain
// `filterFileEntries` exactly (same folding, order, blank-query reference
// semantics); the perf cases pin the property the fix exists for — a
// keystroke filter over a 20k-name corpus stays in milliseconds because the
// folding happened once at index build time, not per keystroke.
import { describe, expect, it } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
import {
  buildNameFilterIndex,
  filterFileEntries,
  filterWithIndex,
} from '../src/renderer/src/home-search'

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

const CORPUS = [
  entry('/docs/Quarterly Report.docx'),
  entry('/sheets/budget.xlsx'),
  entry('/docs/Rapport ÉCOLE.docx'),
  entry('/docs/年度总结.pptx'),
  entry('/home/user/Reports/annual-review.md'),
]

const QUERIES = ['', '   ', 'report', 'REPORT', 'ecole', 'ÉCOLE', '总结', 'annual-review', 'zzz']

describe('filterWithIndex correctness', () => {
  it('matches filterFileEntries exactly across queries', () => {
    const index = buildNameFilterIndex(CORPUS)
    for (const query of QUERIES) {
      expect(filterWithIndex(index, query).map((e) => e.path)).toEqual(
        filterFileEntries(CORPUS, query).map((e) => e.path),
      )
    }
  })

  it('returns the corpus untouched (same reference) for a blank query', () => {
    const index = buildNameFilterIndex(CORPUS)
    expect(filterWithIndex(index, '')).toBe(CORPUS)
    expect(filterWithIndex(index, '   ')).toBe(CORPUS)
  })

  it('never searches the path, only the folded name', () => {
    const index = buildNameFilterIndex([entry('/home/user/Reports/budget.xlsx')])
    expect(filterWithIndex(index, 'reports')).toEqual([])
    expect(filterWithIndex(index, 'budget')).toHaveLength(1)
  })

  it('indexes an empty corpus without exploding', () => {
    const index = buildNameFilterIndex([])
    expect(filterWithIndex(index, 'anything')).toEqual([])
  })
})

describe('filterWithIndex performance at the 20k catalog cap', () => {
  const COUNT = 20_000
  const big = Array.from({ length: COUNT }, (_, i) => {
    const kind = i % 3
    const name =
      kind === 0
        ? `doc-${String(i).padStart(5, '0')}-report.md`
        : kind === 1
          ? `Présentátion ${i}.pptx`
          : `季度报告${i}.xlsx`
    return entry(`/proj/g${Math.floor(i / 500)}/f${i}/${name}`, name)
  })

  it('builds the index over 20k names in bounded time', () => {
    const t0 = performance.now()
    const index = buildNameFilterIndex(big)
    const buildMs = performance.now() - t0
    expect(index.foldedNames).toHaveLength(COUNT)
    // one-time cost; generous bound for slow CI boxes
    expect(buildMs).toBeLessThan(2_000)
  })

  it('keeps every keystroke filter in single-digit/low-double-digit ms', () => {
    const index = buildNameFilterIndex(big)
    const keystrokeQueries = ['d', 'do', 'doc', 'doc-', 'doc-19', 'report', '报告', 'ZZZ']
    const times: number[] = []
    for (const query of keystrokeQueries) {
      const t0 = performance.now()
      const hits = filterWithIndex(index, query)
      times.push(performance.now() - t0)
      // every prefix of the needle matches; the no-hit query still sweeps all
      // 20k folded names, which is exactly the cost being pinned here
      expect(hits.length > 0).toBe(query !== 'ZZZ')
    }
    const worst = Math.max(...times)
    // the old per-keystroke re-fold path cost ~O(n·fold) with a heavy
    // normalize+regex per name; the index sweep is a plain includes pass
    expect(worst).toBeLessThan(100)
  })
})

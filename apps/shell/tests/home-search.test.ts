import { describe, expect, it } from 'vitest'
import type { RecentEntry } from '../src/shared/home-api'
import {
  filterFileEntries,
  foldForSearch,
  isSearchActive,
  mergeFileLists,
} from '../src/renderer/src/home-search'

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

describe('foldForSearch', () => {
  it('lowercases and strips diacritics', () => {
    expect(foldForSearch('ÉCOLE')).toBe('ecole')
    expect(foldForSearch('Ünternehmen')).toBe('unternehmen')
    expect(foldForSearch('résumé')).toBe('resume')
  })

  it('leaves CJK text untouched (no decomposition to fold away)', () => {
    expect(foldForSearch('季度报告')).toBe('季度报告')
    // NFD (not NFKD) does not narrow fullwidth Latin; it only lowercases
    expect(foldForSearch('ＡＢＣ')).toBe('ａｂｃ')
  })
})

describe('isSearchActive', () => {
  it('engages only on non-whitespace queries', () => {
    expect(isSearchActive('')).toBe(false)
    expect(isSearchActive('   \t ')).toBe(false)
    expect(isSearchActive('a')).toBe(true)
    expect(isSearchActive('  报告  ')).toBe(true)
  })
})

describe('filterFileEntries', () => {
  const rows = [
    entry('/docs/Quarterly Report.docx'),
    entry('/sheets/budget.xlsx'),
    entry('/docs/年度总结.pptx'),
  ]

  it('returns the same array untouched for a blank query (no filter)', () => {
    expect(filterFileEntries(rows, '')).toBe(rows)
    expect(filterFileEntries(rows, '   ')).toBe(rows)
  })

  it('keeps entries whose folded name contains the folded query', () => {
    expect(filterFileEntries(rows, 'budget').map((e) => e.path)).toEqual(['/sheets/budget.xlsx'])
    // substring, not prefix
    expect(filterFileEntries(rows, 'arterly').map((e) => e.path)).toEqual([
      '/docs/Quarterly Report.docx',
    ])
    // no match anywhere
    expect(filterFileEntries(rows, 'nope')).toEqual([])
  })

  it('matches case- and accent-insensitively', () => {
    const accented = [entry('/docs/Rapport ÉCOLE.docx')]
    expect(filterFileEntries(accented, 'ecole')).toHaveLength(1)
    expect(filterFileEntries(accented, 'ÉCOLE')).toHaveLength(1)
    expect(filterFileEntries(accented, 'rapport')).toHaveLength(1)
  })

  it('matches CJK names verbatim', () => {
    expect(filterFileEntries(rows, '总结').map((e) => e.path)).toEqual(['/docs/年度总结.pptx'])
    expect(filterFileEntries(rows, '年度总结.pptx')).toHaveLength(1)
  })

  it('searches the file name only, never the path', () => {
    const inReports = [entry('/home/user/Reports/budget.xlsx')]
    expect(filterFileEntries(inReports, 'reports')).toEqual([])
    expect(filterFileEntries(inReports, 'budget')).toHaveLength(1)
  })

  it('preserves input order and trims the query', () => {
    const ordered = [entry('/b/x.md'), entry('/a/x.md'), entry('/c/y.md')]
    expect(filterFileEntries(ordered, '  x.md  ').map((e) => e.path)).toEqual([
      '/b/x.md',
      '/a/x.md',
    ])
  })
})

describe('mergeFileLists', () => {
  it('dedups across recent / starred / project lists, first list wins', () => {
    const recentA = entry('/a/report.docx')
    const starredA = entry('/a/report.docx')
    const projectB = entry('/b/report.docx')
    const merged = mergeFileLists([recentA, projectB], [starredA])
    expect(merged).toHaveLength(2)
    // same path resolves to the earlier list's entry object
    expect(merged[0]).toBe(recentA)
    expect(merged[1]).toBe(projectB)
  })

  it('keeps each list in order: recent rows, then starred-only, then project-only', () => {
    const recent = [entry('/r1.docx'), entry('/r2.docx')]
    const starred = [entry('/r2.docx'), entry('/s1.docx')]
    const project = [entry('/p1.docx'), entry('/s1.docx'), entry('/r1.docx')]
    expect(mergeFileLists(recent, starred, project).map((e) => e.path)).toEqual([
      '/r1.docx',
      '/r2.docx',
      '/s1.docx',
      '/p1.docx',
    ])
  })

  it('handles empty inputs', () => {
    expect(mergeFileLists()).toEqual([])
    expect(mergeFileLists([], [])).toEqual([])
  })
})

/** Merge-source read policy (src/main/merge-source-policy.ts): renderer-named
 *  paths must sit in granted directories before the merge pipeline reads them. */
import { describe, expect, it } from 'vitest'

import {
  checkMergeSourcePath,
  checkMergeSourcePaths,
  expandHomePath,
} from '../src/main/merge-source-policy'

const HOME = '/home/tester'

function deps(mayRead: (path: string) => boolean, exists: (path: string) => boolean) {
  return { homeDir: HOME, mayRead, exists }
}

/** default deps: everything exists, only granted paths readable */
function defaultDeps(granted: (path: string) => boolean) {
  return deps(granted, () => true)
}

describe('expandHomePath', () => {
  it('expands a leading ~/ against the home dir', () => {
    expect(expandHomePath('~/Reports/q1.xlsx', HOME)).toBe(`${HOME}/Reports/q1.xlsx`)
  })

  it('leaves absolute, relative, and bare ~ paths untouched', () => {
    expect(expandHomePath('/tmp/a.csv', HOME)).toBe('/tmp/a.csv')
    expect(expandHomePath('rel/a.csv', HOME)).toBe('rel/a.csv')
    expect(expandHomePath('~', HOME)).toBe('~')
    expect(expandHomePath('~x/a.csv', HOME)).toBe('~x/a.csv')
  })
})

describe('checkMergeSourcePath', () => {
  it('accepts a granted spreadsheet and returns the resolved path', () => {
    expect(
      checkMergeSourcePath(
        '/data/a.xlsx',
        defaultDeps((p) => p === '/data/a.xlsx'),
      ),
    ).toEqual({ resolved: '/data/a.xlsx' })
  })

  it('expands ~/ before the grant check and open', () => {
    const seen: string[] = []
    const result = checkMergeSourcePath(
      '~/Reports/q1.csv',
      deps(
        (p) => (seen.push(p), p === `${HOME}/Reports/q1.csv`),
        () => true,
      ),
    )
    expect(result).toEqual({ resolved: `${HOME}/Reports/q1.csv` })
    expect(seen).toEqual([`${HOME}/Reports/q1.csv`])
  })

  it('rejects unsupported extensions with the offending ext', () => {
    expect(
      checkMergeSourcePath(
        '/data/a.docx',
        defaultDeps(() => true),
      ),
    ).toEqual({
      rejection: { kind: 'ext', ext: 'docx' },
    })
    expect(
      checkMergeSourcePath(
        '/data/noext',
        defaultDeps(() => true),
      ),
    ).toEqual({
      rejection: { kind: 'ext', ext: '' },
    })
  })

  it('rejects ungranted paths before touching the filesystem', () => {
    let existsProbed = false
    const result = checkMergeSourcePath(
      '/etc/passwd.csv',
      deps(
        () => false,
        () => (existsProbed = true),
      ),
    )
    expect(result).toEqual({ rejection: { kind: 'not-granted' } })
    expect(existsProbed).toBe(false)
  })

  it('rejects relative paths (never granted)', () => {
    expect(
      checkMergeSourcePath(
        'rel/a.csv',
        defaultDeps(() => true),
      ),
    ).toEqual({
      rejection: { kind: 'not-granted' },
    })
  })

  it('rejects granted-but-missing files', () => {
    expect(
      checkMergeSourcePath(
        '/data/gone.csv',
        deps(
          () => true,
          () => false,
        ),
      ),
    ).toEqual({
      rejection: { kind: 'missing' },
    })
  })

  it('checks the extension case-insensitively', () => {
    expect(
      checkMergeSourcePath(
        '/data/A.XLSX',
        defaultDeps(() => true),
      ),
    ).toEqual({
      resolved: '/data/A.XLSX',
    })
  })
})

describe('checkMergeSourcePaths', () => {
  it('resolves the whole list when every path is granted', () => {
    const all = (p: string) => p.startsWith('/data/') || p.startsWith(`${HOME}/`)
    expect(
      checkMergeSourcePaths(['/data/a.xlsx', '/data/b.csv', '~/c.xls'], defaultDeps(all)),
    ).toEqual({ resolved: ['/data/a.xlsx', '/data/b.csv', `${HOME}/c.xls`] })
  })

  it('stops at the first rejection without probing later paths', () => {
    const probed: string[] = []
    const result = checkMergeSourcePaths(
      ['/data/a.csv', '/secret/b.csv', '/data/c.csv'],
      deps(
        (p) => (probed.push(p), p.startsWith('/data/')),
        () => true,
      ),
    )
    expect(result).toEqual({ rejection: { kind: 'not-granted' } })
    expect(probed).toEqual(['/data/a.csv', '/secret/b.csv'])
  })
})

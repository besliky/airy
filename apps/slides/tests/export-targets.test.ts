import { describe, expect, it } from 'vitest'

import {
  exportDirInsidePick,
  exportFileMatchesPick,
  isSameExportFile,
  realPathOrDeepestExisting,
  resolveExportImagePaths,
  sanitizeExportBaseName,
} from '../src/main/export-targets'

describe('sanitizeExportBaseName', () => {
  it('accepts and trims ordinary names', () => {
    expect(sanitizeExportBaseName('deck')).toBe('deck')
    expect(sanitizeExportBaseName('  Q4 report ')).toBe('Q4 report')
    expect(sanitizeExportBaseName('a.b.c')).toBe('a.b.c')
  })

  it('rejects traversal and separator payloads', () => {
    expect(sanitizeExportBaseName('')).toBeNull()
    expect(sanitizeExportBaseName('   ')).toBeNull()
    expect(sanitizeExportBaseName('.')).toBeNull()
    expect(sanitizeExportBaseName('..')).toBeNull()
    expect(sanitizeExportBaseName('a/b')).toBeNull()
    expect(sanitizeExportBaseName('a\\b')).toBeNull()
    expect(sanitizeExportBaseName('../../etc/passwd')).toBeNull()
    // windows-style drive +UNC payloads are separators too
    // the colon is rejected now: on NTFS it would address an alternate
    // data stream of the file instead of a regular name
    expect(sanitizeExportBaseName('C:secret')).toBeNull()
  })

  it('rejects control characters', () => {
    expect(sanitizeExportBaseName('a\u0000b')).toBeNull()
    expect(sanitizeExportBaseName('a\u001fb')).toBeNull()
    expect(sanitizeExportBaseName('a\u007fb')).toBeNull()
  })
})

describe('resolveExportImagePaths', () => {
  it('builds padded per-page paths inside the picked directory', () => {
    const paths = resolveExportImagePaths('/tmp/out', 'deck', 2)
    expect(paths).toEqual(['/tmp/out/deck-01.png', '/tmp/out/deck-02.png'])
  })

  it('widens the padding to three digits for 100+ pages', () => {
    const paths = resolveExportImagePaths('/tmp/out', 'deck', 100)
    expect(paths?.[0]).toBe('/tmp/out/deck-001.png')
    expect(paths?.[99]).toBe('/tmp/out/deck-100.png')
    expect(paths?.length).toBe(100)
  })

  it('refuses unsafe base names instead of escaping the directory', () => {
    expect(resolveExportImagePaths('/tmp/out', '../evil', 1)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', 'a/b', 1)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', '..', 1)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', '', 1)).toBeNull()
  })

  it('refuses nonsensical page counts', () => {
    expect(resolveExportImagePaths('/tmp/out', 'deck', 0)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', 'deck', -3)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', 'deck', Number.NaN)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', 'deck', 1e9)).toBeNull()
  })

  it('keeps containment on win32-style inputs regardless of host separators', () => {
    const paths = resolveExportImagePaths('C:\\Users\\me\\Exports', 'deck', 1, 'win32')
    // the host join may keep posix separators; win32 normalization must still land in the picked dir
    expect(paths?.length).toBe(1)
    expect(isSameExportFile('C:\\Users\\me\\Exports\\deck-01.png', paths![0], 'win32')).toBe(true)
    expect(isSameExportFile('C:\\Users\\me\\evil.png', paths![0], 'win32')).toBe(false)
  })

  it('resolves .jpg paths for the jpeg format (PAR-316)', () => {
    expect(resolveExportImagePaths('/tmp/out', 'deck', 2, 'linux', 'jpg')).toEqual([
      '/tmp/out/deck-01.jpg',
      '/tmp/out/deck-02.jpg',
    ])
    // png stays the explicit default
    expect(resolveExportImagePaths('/tmp/out', 'deck', 1, 'linux', 'png')).toEqual([
      '/tmp/out/deck-01.png',
    ])
  })

  it('refuses unsupported extensions instead of writing surprise file types', () => {
    expect(resolveExportImagePaths('/tmp/out', 'deck', 1, 'linux', 'exe' as never)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', 'deck', 1, 'linux', 'JPG' as never)).toBeNull()
    expect(resolveExportImagePaths('/tmp/out', '../evil', 1, 'linux', 'jpg')).toBeNull()
  })
})

describe('isSameExportFile', () => {
  it('matches identical and whitespace-padded paths', () => {
    expect(isSameExportFile('/tmp/a.pdf', '/tmp/a.pdf')).toBe(true)
    expect(isSameExportFile('/tmp/a.pdf', '  /tmp/a.pdf ')).toBe(true)
  })

  it('rejects different paths and empty input', () => {
    expect(isSameExportFile('/tmp/a.pdf', '/tmp/b.pdf')).toBe(false)
    expect(isSameExportFile('/tmp/a.pdf', '/tmp/a.pdf.bak')).toBe(false)
    expect(isSameExportFile('', '/tmp/a.pdf')).toBe(false)
    expect(isSameExportFile('/tmp/a.pdf', '')).toBe(false)
  })

  it('normalizes dot segments and (on win32) case', () => {
    expect(isSameExportFile('/tmp/x/../a.pdf', '/tmp/a.pdf')).toBe(true)
    expect(isSameExportFile('C:\\Users\\me\\a.pdf', 'c:\\users\\ME\\a.pdf', 'win32')).toBe(true)
    expect(isSameExportFile('C:\\Users\\me\\a.pdf', 'C:\\Users\\me\\B.pdf', 'win32')).toBe(false)
  })
})

describe('sanitizeExportBaseName (windows-reserved characters)', () => {
  it('rejects the NTFS alternate-data-stream colon', () => {
    expect(sanitizeExportBaseName('name:hidden')).toBeNull()
    expect(sanitizeExportBaseName('a:b-01')).toBeNull()
  })

  it('rejects the other Windows-reserved filename characters', () => {
    for (const name of ['a*b', 'a?b', 'a"b', 'a<b', 'a>b', 'a|b']) {
      expect(sanitizeExportBaseName(name), name).toBeNull()
    }
  })

  it('still accepts ordinary names and dots inside the name', () => {
    expect(sanitizeExportBaseName('  deck final  ')).toBe('deck final')
    expect(sanitizeExportBaseName('deck.v2')).toBe('deck.v2')
  })
})

describe('realPathOrDeepestExisting', () => {
  /** fake filesystem: only the listed real paths exist; symlinks re-map */
  function fakeRealpath(real: Record<string, string>): (p: string) => string {
    return (p: string) => {
      const target = real[p]
      if (target === undefined) throw new Error('ENOENT')
      return target
    }
  }

  it('returns the realpath for an existing path', () => {
    const realpath = fakeRealpath({ '/pick/sub': '/data/sub' })
    expect(realPathOrDeepestExisting('/pick/sub', realpath)).toBe('/data/sub')
  })

  it('resolves a missing tail through its deepest existing ancestor', () => {
    const realpath = fakeRealpath({ '/pick': '/data' })
    expect(realPathOrDeepestExisting('/pick/new/deck.pdf', realpath)).toBe('/data/new/deck.pdf')
  })

  it('falls back to the lexical spelling when nothing exists', () => {
    const realpath = fakeRealpath({})
    expect(realPathOrDeepestExisting('/gone/x', realpath)).toBe('/gone/x')
  })
})

describe('exportDirInsidePick (physical containment)', () => {
  /** realpaths: the pick and a subdirectory exist; a swapped pick is a symlink */
  const realpathOf = (p: string): string => {
    const map: Record<string, string> = {
      '/pick': '/data/export',
      '/pick/sub': '/data/export/sub',
    }
    const target = map[p]
    if (target === undefined) throw new Error('ENOENT')
    return target
  }

  it('accepts the picked dir itself and existing/new subdirectories', () => {
    expect(exportDirInsidePick('/data/export', '/pick', realpathOf)).toBe(true)
    expect(exportDirInsidePick('/data/export', '/pick/sub', realpathOf)).toBe(true)
    // fresh subdirectory resolves through the existing pick
    expect(exportDirInsidePick('/data/export', '/pick/brand-new', realpathOf)).toBe(true)
  })

  it('rejects a swapped pick: the stored realpath no longer matches what the name resolves to', () => {
    // after the swap, /elsewhere is itself real and /pick now maps outside
    const swapped = (p: string): string => {
      const map: Record<string, string> = {
        '/pick': '/elsewhere/attacker',
        '/pick/sub': '/elsewhere/attacker/sub',
      }
      const target = map[p]
      if (target === undefined) throw new Error('ENOENT')
      return target
    }
    expect(exportDirInsidePick('/data/export', '/pick', swapped)).toBe(false)
    expect(exportDirInsidePick('/data/export', '/pick/sub', swapped)).toBe(false)
  })

  it('rejects sibling and outside dirs outright', () => {
    expect(exportDirInsidePick('/data/export', '/other', realpathOf)).toBe(false)
    expect(exportDirInsidePick('/data/export', '/data/export-evil', realpathOf)).toBe(false)
  })
})

describe('exportFileMatchesPick (physical file match)', () => {
  const realpathOf = (p: string): string => {
    const map: Record<string, string> = { '/data': '/data' }
    const target = map[p]
    if (target === undefined) throw new Error('ENOENT')
    return target
  }

  it('matches an existing file and a not-yet-created target in the picked folder', () => {
    expect(exportFileMatchesPick('/data/deck.pdf', '/data/deck.pdf', realpathOf)).toBe(true)
    expect(exportFileMatchesPick('/data/deck.pdf', '/data/deck.pdf', realpathOf)).toBe(true)
  })

  it('rejects a different file and a swapped directory component', () => {
    expect(exportFileMatchesPick('/data/deck.pdf', '/data/other.pdf', realpathOf)).toBe(false)
    const swapped = (p: string): string => {
      const map: Record<string, string> = { '/data': '/elsewhere' }
      const target = map[p]
      if (target === undefined) throw new Error('ENOENT')
      return target
    }
    expect(exportFileMatchesPick('/data/deck.pdf', '/data/deck.pdf', swapped)).toBe(false)
  })
})

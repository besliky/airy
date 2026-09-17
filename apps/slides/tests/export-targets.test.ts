import { describe, expect, it } from 'vitest'

import {
  isSameExportFile,
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
    expect(sanitizeExportBaseName('C:secret')).toBe('C:secret') // bare colon stays a legal ntfs name char
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

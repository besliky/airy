import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, sep } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

/**
 * Untitled staging (src/main/untitled-staging.ts): blank sheet/pdf files are
 * staged under userData until their first explicit save; these tests cover
 * the pure containment/listing/purge decisions that keep the staging dir from
 * leaking files or deleting anything outside it.
 */

let S: typeof import('../src/main/untitled-staging')
let scratch: string
let staging: string

beforeEach(async () => {
  S = await import('../src/main/untitled-staging')
  scratch = mkdtempSync(join(tmpdir(), 'airy-untitled-staging-'))
  staging = S.untitledStagingDir(scratch)
  mkdirSync(staging, { recursive: true })
})

afterEach(() => {
  rmSync(scratch, { recursive: true, force: true })
})

describe('untitledStagingDir', () => {
  it('lives under userData', () => {
    expect(S.untitledStagingDir('/user/data')).toBe(join('/user/data', 'untitled-staging'))
  })
})

describe('isInsideDirectory', () => {
  it('accepts files directly in the directory and in subfolders', () => {
    expect(S.isInsideDirectory(staging, join(staging, 'Untitled Spreadsheet.xlsx'))).toBe(true)
    expect(S.isInsideDirectory(staging, join(staging, 'sub', 'x.pdf'))).toBe(true)
  })

  it('rejects sibling prefixes and outside paths', () => {
    // the directory itself counts as inside (harmless: removing it fails safely)
    expect(S.isInsideDirectory(staging, staging)).toBe(true)
    expect(S.isInsideDirectory(staging, staging + '-other/x.pdf')).toBe(false)
    // join normalizes the .. away, so the escaped path no longer matches
    expect(S.isInsideDirectory(staging, join(staging, '..', 'elsewhere.pdf'))).toBe(false)
  })
})

describe('listStagedFiles', () => {
  it('lists only staged extensions', () => {
    writeFileSync(join(staging, 'a.xlsx'), '')
    writeFileSync(join(staging, 'b.pdf'), '')
    writeFileSync(join(staging, 'c.txt'), '')
    writeFileSync(join(staging, 'd'), '')
    const listed = S.listStagedFiles(staging).map((p) => p.split(sep).pop())
    expect(listed).toEqual(['a.xlsx', 'b.pdf'])
  })

  it('returns an empty list when the directory does not exist', () => {
    expect(S.listStagedFiles(join(scratch, 'absent'))).toEqual([])
  })
})

describe('removeStagedFile', () => {
  it('deletes a file inside the staging dir', () => {
    const file = join(staging, 'a.pdf')
    writeFileSync(file, '')
    expect(S.removeStagedFile(staging, file)).toBe(true)
    expect(S.listStagedFiles(staging)).toEqual([])
  })

  it('refuses paths outside the staging dir (already-rebound tabs)', () => {
    const outside = join(scratch, 'real.pdf')
    writeFileSync(outside, '')
    expect(S.removeStagedFile(staging, outside)).toBe(false)
    expect(S.listStagedFiles(staging)).toEqual([outside].filter(() => false)) // outside was never staged
  })

  it('is idempotent for already-deleted staged files', () => {
    expect(S.removeStagedFile(staging, join(staging, 'gone.xlsx'))).toBe(true)
  })
})

describe('orphanedStagedFiles', () => {
  it('flags staged files no open tab owns', () => {
    const staged = ['/s/a.xlsx', '/s/b.pdf', '/s/c.xlsx']
    const open = ['/s/a.xlsx']
    expect(S.orphanedStagedFiles(staged, open)).toEqual(['/s/b.pdf', '/s/c.xlsx'])
  })

  it('keeps files an open tab still shows (crash survivors restored)', () => {
    const staged = ['/s/a.xlsx']
    expect(S.orphanedStagedFiles(staged, ['/s/a.xlsx'])).toEqual([])
  })
})

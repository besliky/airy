import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeAll, afterAll, describe, expect, it, vi } from 'vitest'

import {
  _setRenameRetryDelayForTests,
  atomicWriteFile,
  looksLikeZip,
  renameDurably,
} from '../src/atomic-write'

// PERF-1101: the retry ladder is pinned by call counts below; its real
// 50..800 ms pacing would add ~3 s of sleeps to this suite, so collapse it.
beforeAll(() => _setRenameRetryDelayForTests(0))
afterAll(() => _setRenameRetryDelayForTests(null))

// fsync order is the durability contract: temp-file sync must precede the
// rename, and the directory sync (POSIX only) must follow it.
const { order, setInPlaceWriteError, takeInPlaceWriteError } = vi.hoisted(() => {
  let inPlaceError: Error | null = null
  return {
    order: [] as string[],
    setInPlaceWriteError: (error: Error | null) => {
      inPlaceError = error
    },
    // the non-atomic in-place fallback rides a file handle (not the
    // module-level writeFile) — tests inject its failure there
    takeInPlaceWriteError: () => {
      const error = inPlaceError
      inPlaceError = null
      return error
    },
  }
})

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    rename: vi.fn((...args: Parameters<typeof actual.rename>) => {
      order.push('rename')
      return actual.rename(...args)
    }),
    open: vi.fn(async (...args: Parameters<typeof actual.open>) => {
      const handle = await actual.open(...args)
      const path = args[0]
      const isTemp = typeof path === 'string' && path.includes('.tmp')
      // syncDirectory is the only reader: it opens directories with 'r',
      // while every file write (temp and in-place fallback) opens with 'w'
      const isDirectory = args[1] === 'r'
      const realSync = handle.sync.bind(handle)
      handle.sync = async () => {
        order.push(isDirectory ? 'dir-sync' : 'file-sync')
        await realSync()
      }
      const realWriteFile = handle.writeFile.bind(handle)
      handle.writeFile = async (...writeArgs: Parameters<typeof handle.writeFile>) => {
        const error = isTemp ? null : takeInPlaceWriteError()
        if (error) throw error
        return realWriteFile(...writeArgs)
      }
      return handle
    }),
  }
})

let dir = ''

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = ''
  vi.mocked(rename).mockClear()
  setInPlaceWriteError(null)
  order.length = 0
})

const epermError = () =>
  Object.assign(new Error('EPERM: operation not permitted'), { code: 'EPERM' })

describe('atomicWriteFile', () => {
  it('replaces the target and leaves no temp file behind', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')

    await atomicWriteFile(target, Buffer.from('new'))

    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('keeps the original intact and cleans the temp file when the write fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    const bad = { byteLength: -1 } as unknown as Buffer

    await expect(atomicWriteFile(target, bad)).rejects.toThrow()

    expect(readFileSync(target, 'utf-8')).toBe('old')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  // Windows: Defender/indexer briefly locks the target and rename throws EPERM
  it('retries transient Windows rename locks and still lands atomically', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(epermError()).mockRejectedValueOnce(epermError())

    await atomicWriteFile(target, Buffer.from('new'))

    expect(vi.mocked(rename)).toHaveBeenCalledTimes(3)
    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('falls back to an in-place write when the rename stays locked', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    for (let i = 0; i < 5; i++) vi.mocked(rename).mockRejectedValueOnce(epermError())

    await atomicWriteFile(target, Buffer.from('new'))

    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  // BUG-1222: the double-failure path (rename stayed locked AND the in-place
  // fallback failed) used to skip both unlinks and orphan the temp next to
  // the target; it is cleaned either way now, without masking the failure.
  it('cleans the temp file when the fallback write fails too (BUG-1222)', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    for (let attempt = 0; attempt < 5; attempt += 1) {
      vi.mocked(rename).mockRejectedValueOnce(epermError())
    }
    // the temp write rides a file handle first; the next non-temp handle
    // write is the non-atomic fallback itself
    setInPlaceWriteError(Object.assign(new Error('EIO: fallback write failed'), { code: 'EIO' }))

    await expect(atomicWriteFile(target, Buffer.from('new'))).rejects.toThrow(
      'fallback write failed',
    )

    // no temp orphan remains next to the target (the in-place fallback's
    // open('w') truncating the target on its own failure is that fallback's
    // pre-existing hazard, unchanged by this fix and out of its scope)
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  // BUG-1203: the fallback writes through a file handle and fsyncs it —
  // in-place writes change no dirent, but the bytes must still be durable.
  it('fsyncs the in-place fallback write instead of a bare writeFile', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    for (let attempt = 0; attempt < 5; attempt += 1) {
      vi.mocked(rename).mockRejectedValueOnce(epermError())
    }

    await atomicWriteFile(target, Buffer.from('new'))

    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(vi.mocked(rename)).toHaveBeenCalledTimes(5)
    // the temp sync and the flushed in-place write; no directory sync — an
    // in-place write changes no directory entry (mockRejectedValueOnce
    // replaces the mock body, so refused renames don't reach `order`)
    expect(order).toEqual(['file-sync', 'file-sync'])
  })

  it('uses distinct temp files for concurrent writes to one target', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')

    await Promise.all([
      atomicWriteFile(target, Buffer.from('first')),
      atomicWriteFile(target, Buffer.from('second')),
    ])

    expect(['first', 'second']).toContain(readFileSync(target, 'utf-8'))
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  // Durability (BUG-710): an un-synced rename can surface as an empty or
  // truncated file after power loss; the bytes must reach the disk first.
  it('fsyncs the temp file before the rename and the directory after it', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')

    await atomicWriteFile(target, Buffer.from('new'))

    expect(readFileSync(target, 'utf-8')).toBe('new')
    if (process.platform === 'win32') {
      // no directory handle is opened on Windows (fsync on a dir is EPERM)
      expect(order).toEqual(['file-sync', 'rename'])
    } else {
      expect(order).toEqual(['file-sync', 'rename', 'dir-sync'])
    }
  })
})

describe('renameDurably', () => {
  // BUG-1203: the sheets xlsx promote/write paths share this helper, so its
  // protocol (retry + POSIX dir-fsync after the rename) is pinned here once.
  it('renames and fsyncs the parent directory afterwards (POSIX)', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    writeFileSync(temporary, 'new')

    await renameDurably(temporary, target)

    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['book.xlsx'])
    if (process.platform === 'win32') {
      expect(order).toEqual(['rename'])
    } else {
      expect(order).toEqual(['rename', 'dir-sync'])
    }
  })

  it('rethrows the original retryable code after exhausting the retries', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    writeFileSync(temporary, 'new')
    vi.mocked(rename).mockRejectedValue(epermError())

    await expect(renameDurably(temporary, target)).rejects.toMatchObject({ code: 'EPERM' })
    // initial attempt + RENAME_RETRIES retries
    expect(vi.mocked(rename)).toHaveBeenCalledTimes(5)
    // the surviving temp is the caller's fallback input, not ours to delete
    expect(readFileSync(temporary, 'utf-8')).toBe('new')
  })
})

describe('looksLikeZip', () => {
  it('accepts a zip local-file header and rejects truncated/garbage bytes', () => {
    expect(looksLikeZip(Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00]))).toBe(true)
    expect(looksLikeZip(Buffer.from([0x50, 0x4b]))).toBe(false)
    expect(looksLikeZip(Buffer.from('not a zip'))).toBe(false)
    expect(looksLikeZip(Buffer.alloc(0))).toBe(false)
  })
})

import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { rename, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

import { atomicWriteFile, looksLikeZip } from '../src/atomic-write'

// fsync order is the durability contract: temp-file sync must precede the
// rename, and the directory sync (POSIX only) must follow it.
const { order } = vi.hoisted(() => ({ order: [] as string[] }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    rename: vi.fn((...args: Parameters<typeof actual.rename>) => {
      order.push('rename')
      return actual.rename(...args)
    }),
    writeFile: vi.fn(actual.writeFile),
    open: vi.fn(async (...args: Parameters<typeof actual.open>) => {
      const handle = await actual.open(...args)
      const path = args[0]
      const isDirectory = typeof path === 'string' && !path.includes('.tmp')
      const realSync = handle.sync.bind(handle)
      handle.sync = async () => {
        order.push(isDirectory ? 'dir-sync' : 'file-sync')
        await realSync()
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
  vi.mocked(writeFile).mockClear()
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

  it('preserves the completed temp file when the fallback write fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    for (let attempt = 0; attempt < 5; attempt += 1) {
      vi.mocked(rename).mockRejectedValueOnce(epermError())
    }
    const fallbackError = Object.assign(new Error('EIO: fallback write failed'), { code: 'EIO' })
    // the temp write rides a file handle now, so the first module-level
    // writeFile call is the non-atomic fallback itself
    vi.mocked(writeFile).mockRejectedValueOnce(fallbackError)

    await expect(atomicWriteFile(target, Buffer.from('new'))).rejects.toThrow(
      'fallback write failed',
    )

    const files = readdirSync(dir)
    expect(files).toContain('a.docx')
    const temp = files.find((file) => file !== 'a.docx')
    expect(temp).toBeDefined()
    expect(readFileSync(join(dir, temp!), 'utf-8')).toBe('new')
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

describe('looksLikeZip', () => {
  it('accepts a zip local-file header and rejects truncated/garbage bytes', () => {
    expect(looksLikeZip(Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00]))).toBe(true)
    expect(looksLikeZip(Buffer.from([0x50, 0x4b]))).toBe(false)
    expect(looksLikeZip(Buffer.from('not a zip'))).toBe(false)
    expect(looksLikeZip(Buffer.alloc(0))).toBe(false)
  })
})

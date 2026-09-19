import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'

import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

import { syncFileBestEffort, writeXlsxAtomically } from '../src/gateway/xlsx-gateway'

// The only fs/promises directory reads in this suite come from the shared
// promote helper's POSIX dir-fsync after the rename (BUG-1203).
const { directoryOpens } = vi.hoisted(() => ({ directoryOpens: [] as string[] }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    open: vi.fn(async (...args: Parameters<typeof actual.open>) => {
      const [path, flags] = args
      if (flags === 'r') directoryOpens.push(String(path))
      return actual.open(...args)
    }),
  }
})

beforeEach(() => {
  directoryOpens.length = 0
})

describe('syncFileBestEffort', () => {
  let directory: string

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'xlsx-atomic-sync-'))
  })

  afterAll(async () => {
    await rm(directory, { recursive: true, force: true })
  })

  it('flushes a writable file without altering it', async () => {
    const path = join(directory, 'plain.bin')
    await writeFile(path, 'payload')
    await syncFileBestEffort(path)
    expect(await readFile(path, 'utf8')).toBe('payload')
  })

  it('tolerates a file it cannot reopen for writing (cloud-sync/AV lock shape)', async () => {
    const path = join(directory, 'readonly.bin')
    await writeFile(path, 'payload')
    await chmod(path, 0o444)
    await expect(syncFileBestEffort(path)).resolves.toBeUndefined()
  })

  it('still surfaces a missing file', async () => {
    await expect(syncFileBestEffort(join(directory, 'absent.bin'))).rejects.toMatchObject({
      code: 'ENOENT',
    })
  })
})

describe('writeXlsxAtomically', () => {
  it('writes through a flushed temp file and renames it into place', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-atomic-write-'))
    try {
      const path = join(directory, 'book.xlsx')
      await writeXlsxAtomically(path, Buffer.from('bytes'))
      expect(await readFile(path, 'utf8')).toBe('bytes')
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })

  // BUG-1203: the rename rides the shared durability helper, so the parent
  // directory is fsynced on POSIX after it (skipped on Windows, where a
  // directory cannot be opened for fsync).
  it('directory-fsyncs after the rename via the shared promote helper (POSIX)', async () => {
    if (process.platform === 'win32') return
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-atomic-write-'))
    try {
      const path = join(directory, 'book.xlsx')
      await writeXlsxAtomically(path, Buffer.from('bytes'))
      expect(await readFile(path, 'utf8')).toBe('bytes')
      expect(directoryOpens).toEqual([dirname(path)])
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })
})

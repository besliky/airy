/**
 * Saving promotes a same-directory temp file over the target with rename.
 * Windows AV/indexer/cloud-sync locks make that rename fail transiently with
 * EPERM/EACCES/EBUSY (alpha: "EPERM: operation not permitted, rename
 * .tmp.xlsx → …"), so the promotion retries, falls back to an in-place copy,
 * and surfaces a stable, localizable message when the target stays locked.
 */
import {
  chmod,
  copyFile,
  link,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  stat,
  truncate,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, beforeAll, afterAll, describe, expect, it, vi } from 'vitest'
import { _setRenameRetryDelayForTests } from '@airy-office/electron-utils/atomic-write'
import { promoteFileAtomically, promoteFileExclusively } from '../src/gateway/xlsx-package-io'

// PERF-1101: these locks force the shared renameDurably retry ladder; its
// real 50..800 ms pacing would add ~3 s of sleeps, so collapse it (the retry
// count and the fallback routing stay pinned by the assertions below).
beforeAll(() => _setRenameRetryDelayForTests(0))
afterAll(() => _setRenameRetryDelayForTests(null))

// Directory handles opened for reading are the shared helper's POSIX
// dir-fsync after the rename (BUG-1203) — nothing else in this suite opens
// directories through fs/promises.
const { directoryOpens } = vi.hoisted(() => ({ directoryOpens: [] as string[] }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    copyFile: vi.fn(actual.copyFile),
    link: vi.fn(actual.link),
    open: vi.fn(async (...args: Parameters<typeof actual.open>) => {
      const [path, flags] = args
      if (flags === 'r') directoryOpens.push(String(path))
      return actual.open(...args)
    }),
  }
})
const copyFileMock = vi.mocked(copyFile)
const linkMock = vi.mocked(link)

const scratches: string[] = []
const actualCopyFile = copyFileMock.getMockImplementation()!

async function scratchDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'airy-promote-test-'))
  scratches.push(dir)
  return dir
}

const actualLink = linkMock.getMockImplementation()!

afterEach(async () => {
  copyFileMock.mockReset()
  copyFileMock.mockImplementation(actualCopyFile)
  linkMock.mockReset()
  linkMock.mockImplementation(actualLink)
  directoryOpens.length = 0
  for (const dir of scratches.splice(0)) {
    await chmod(dir, 0o755).catch(() => {})
    for (const sub of ['locked']) await chmod(join(dir, sub), 0o755).catch(() => {})
    await rm(dir, { recursive: true, force: true })
  }
})

describe('promoteFileAtomically', () => {
  it('replaces the target and removes the temp file', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    await promoteFileAtomically(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  // BUG-1203: the workbook save promote rides the shared durability helper —
  // on POSIX the parent directory is fsynced after the rename so the new
  // dirent survives power loss (Windows skips it: dir fsync is EPERM there).
  it('fsyncs the parent directory after the rename lands (POSIX)', async () => {
    if (process.platform === 'win32') return
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')

    await promoteFileAtomically(temporary, target)

    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    expect(directoryOpens).toEqual([dir])
  })

  it('falls back to an in-place copy when only the rename is blocked', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    // a read-only directory rejects the rename with the same retryable
    // EACCES a Windows lock produces, while the target file stays writable
    await chmod(locked, 0o555)
    await promoteFileAtomically(temporary, target)
    await chmod(locked, 0o755)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
  }, 15_000)

  it('reports a persistently locked target with a stable localizable message', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    // rename AND in-place copy both refused — the Excel-holds-the-file case
    await chmod(target, 0o444)
    await chmod(locked, 0o555)
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    await chmod(locked, 0o755)
    await chmod(target, 0o644)
    // the finished bytes survive the failure for the caller's cleanup/retry
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
  }, 15_000)

  it('restores the target when the in-place copy dies after truncating it', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    await chmod(locked, 0o555)
    // the backup copy runs for real; the copy over the target truncates it
    // and then fails the way a lock acquired mid-write does
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) !== temporary) return actualCopyFile(src, dest, mode)
      await truncate(String(dest))
      throw Object.assign(new Error('EBUSY: resource busy'), { code: 'EBUSY' })
    })
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    await chmod(locked, 0o755)
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
  }, 15_000)

  it('names the surviving backup when the target cannot be restored either', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    await chmod(locked, 0o555)
    const busy = () => Object.assign(new Error('EBUSY: resource busy'), { code: 'EBUSY' })
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) === temporary) {
        await truncate(String(dest))
        throw busy()
      }
      if (String(dest) === target) throw busy()
      return actualCopyFile(src, dest, mode)
    })
    const failure = await promoteFileAtomically(temporary, target).catch((error: Error) => error)
    await chmod(locked, 0o755)
    expect(failure?.message).toContain('preserved at: ')
    const survivor = failure!.message.split('preserved at: ')[1] ?? ''
    // the read-only directory refuses the recovered copy, so the tmp backup stays
    expect(survivor.startsWith(tmpdir())).toBe(true)
    expect(basename(survivor)).toMatch(/^book\.recovered-[0-9a-f-]+\.xlsx$/)
    expect(await readFile(survivor, 'utf8')).toBe('old-bytes')
    await rm(survivor, { force: true })
  }, 15_000)

  it('propagates non-retryable errors untouched', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.missing.tmp.xlsx')
    await expect(promoteFileAtomically(temporary, join(dir, 'book.xlsx'))).rejects.toMatchObject({
      code: 'ENOENT',
    })
  })
})

describe('promoteFileExclusively', () => {
  it('creates the target and removes the temp file', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await promoteFileExclusively(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('refuses a target that appeared after the caller check (EEXIST) and cleans the temp', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    // a writer created the target inside the guard/write window: the link
    // fails with EEXIST atomically and the save aborts instead of replacing
    await writeFile(target, 'concurrent bytes')
    await expect(promoteFileExclusively(temporary, target)).rejects.toThrow(
      'The save target already exists',
    )
    expect(await readFile(target, 'utf8')).toBe('concurrent bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('propagates unrelated link errors untouched', async () => {
    const dir = await scratchDir()
    await expect(
      promoteFileExclusively(join(dir, '.missing.tmp.xlsx'), join(dir, 'book.xlsx')),
    ).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('EPERM from link (exFAT/FAT/network) with an existing target still refuses', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'concurrent bytes')
    linkMock.mockImplementation(async () => {
      throw Object.assign(new Error('EPERM: operation not permitted, link'), { code: 'EPERM' })
    })
    await expect(promoteFileExclusively(temporary, target)).rejects.toThrow(
      'The save target already exists',
    )
    expect(await readFile(target, 'utf8')).toBe('concurrent bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('EPERM from link with a missing target falls back to the atomic rename promote', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    linkMock.mockImplementation(async () => {
      throw Object.assign(new Error('EPERM: operation not permitted, link'), { code: 'EPERM' })
    })
    await promoteFileExclusively(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('EACCES from link behaves like EPERM (fallback on missing target)', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    linkMock.mockImplementation(async () => {
      throw Object.assign(new Error('EACCES: permission denied, link'), { code: 'EACCES' })
    })
    await promoteFileExclusively(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
  })
})

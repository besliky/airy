/**
 * A CSV in-place save writes back into the user's original .csv — usually the
 * only copy. The write-back must stage the bytes and rename them into place:
 * a plain writeFile truncates the target before the data lands, so a crash or
 * disk error mid-write destroys the file. These tests pin the saved-file
 * contents (UTF-8 BOM + full payload) and the failure shape (an error
 * mid-write leaves the original file complete, no truncated or leftover
 * staging files).
 */
import { readFileSync } from 'node:fs'
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it, vi } from 'vitest'

// The helper pulls atomicWriteFile from the electron-utils barrel; two of its
// modules bind electron values at import time and never call them here.
vi.mock('electron', () => ({ ipcRenderer: {}, webUtils: {}, shell: {} }))

// atomicWriteFile stages through a file handle (write + fsync + rename), so
// the mid-write death is simulated on the handle's writeFile.
const { staging } = vi.hoisted(() => ({ staging: { failOnce: false } }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    writeFile: vi.fn(actual.writeFile),
    open: vi.fn(async (...args: Parameters<typeof actual.open>) => {
      const handle = await actual.open(...args)
      const realWriteFile = handle.writeFile.bind(handle)
      handle.writeFile = (async (
        data: string | NodeJS.ArrayBufferView,
        options?: Parameters<typeof realWriteFile>[1],
      ) => {
        if (!staging.failOnce) return realWriteFile(data, options)
        staging.failOnce = false
        // land a few bytes, then die the way a full disk or a killed
        // process does
        const staged =
          typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data as Uint8Array)
        await realWriteFile(staged.subarray(0, 3), options)
        throw Object.assign(new Error('ENOSPC: no space left on device, write'), {
          code: 'ENOSPC',
        })
      }) as typeof handle.writeFile
      return handle
    }),
  }
})
const writeFileMock = vi.mocked(writeFile)
const actualWriteFile = writeFileMock.getMockImplementation()!

import { csvBytesWithBom, writeCsvBackAtomic } from '../src/main/csv-save-back'

const here = dirname(fileURLToPath(import.meta.url))
const scratches: string[] = []

async function scratchDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'airy-csv-save-back-'))
  scratches.push(dir)
  return dir
}

afterEach(async () => {
  writeFileMock.mockReset()
  writeFileMock.mockImplementation(actualWriteFile)
  staging.failOnce = false
  for (const dir of scratches.splice(0)) {
    await rm(dir, { recursive: true, force: true })
  }
})

describe('writeCsvBackAtomic', () => {
  it('replaces the target with the complete BOM-prefixed content', async () => {
    const dir = await scratchDir()
    const path = join(dir, 'data.csv')
    await writeFile(path, 'old,bytes\r\n', 'utf8')
    await writeCsvBackAtomic(path, 'name,note\r\nAda,"says ""ok"""\r\n')
    const onDisk = await readFile(path)
    expect([...onDisk.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    expect(onDisk.equals(csvBytesWithBom('name,note\r\nAda,"says ""ok"""\r\n'))).toBe(true)
    // the staged temp is renamed into place, not left beside the save
    expect((await readdir(dir)).sort()).toEqual(['data.csv'])
  })

  it('a write that dies mid-flight leaves the original file complete', async () => {
    const dir = await scratchDir()
    const path = join(dir, 'data.csv')
    const before = csvBytesWithBom('a,b\r\n1,2\r\n')
    await writeFile(path, before)
    // The staging write lands a few bytes then dies the way a full disk or a
    // killed process does. The old direct writeFile(path, …) had already
    // truncated the user's csv by this point.
    staging.failOnce = true
    await expect(writeCsvBackAtomic(path, 'a,b\r\n3,4\r\n')).rejects.toMatchObject({
      code: 'ENOSPC',
    })
    expect((await readFile(path)).equals(before)).toBe(true)
    expect((await readdir(dir)).sort()).toEqual(['data.csv'])
  })
})

describe('CSV save-back wiring', () => {
  it('the save handler routes the in-place CSV write-back through writeCsvBackAtomic', () => {
    const mainSrc = readFileSync(join(here, '../src/main/sheets-main.ts'), 'utf8')
    expect(mainSrc).toContain('writeCsvBackAtomic(session.csvSourcePath, request.csvContent)')
  })
})

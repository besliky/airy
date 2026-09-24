/**
 * OBS-1670: a kill -9 mid atomicWriteFile bypasses every error-path unlink,
 * so the dot-prefixed `.<target>.<12 hex>.tmp` save temp survives next to the
 * user's document forever — a directory nothing app-owned ever revisits. The
 * fix: each fresh save (savePdfToPath) sweeps the target directory's expired
 * temps first. Age alone decides, so an in-flight save's temp is never a
 * candidate. The sweep lives in save-temp-sweep.ts (one pattern owner next to
 * the construction it mirrors in electron-utils/atomic-write.ts) and the
 * save-pdf.ts wiring is pinned source-style, as in the slides OBS-1658 sweep.
 */
import { existsSync, readFileSync } from 'node:fs'
import { mkdir, mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { randomBytes } from 'node:crypto'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'

import { SAVE_TEMP_TTL_MS, sweepStaleSaveTemps } from '../src/main/save-temp-sweep'
import { savePdfToPath } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function tempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'airy-pdf-save-tmp-test-'))
  roots.push(dir)
  return dir
}

/** A temp name of the exact shape atomicWriteFile produces before its rename. */
const orphanName = (target: string) => `.${target}.${'ab12cd34ef56'}.tmp`

async function makePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([612, 792])
  return doc.save({ useObjectStreams: false })
}

const request = (path: string, over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path,
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

describe('stale PDF-save temp sweep (OBS-1670)', () => {
  it('removes an expired kill -9 leftover temp next to the save target', async () => {
    const dir = await tempDir()
    const leftover = join(dir, orphanName('report.pdf'))
    await writeFile(leftover, 'orphaned bytes')
    const now = Date.now()
    const stale = new Date(now - SAVE_TEMP_TTL_MS - 60_000)
    await utimes(leftover, stale, stale)

    const removed = await sweepStaleSaveTemps(dir, now)

    expect(removed).toEqual([leftover])
    expect(existsSync(leftover)).toBe(false)
  })

  it('never sweeps a live save temp: anything within the TTL window stays', async () => {
    const dir = await tempDir()
    const now = Date.now()
    // a temp written seconds ago (in-progress save)
    const live = join(dir, orphanName('report.pdf'))
    await writeFile(live, 'writing…')
    const fresh = new Date(now - 5_000)
    await utimes(live, fresh, fresh)
    // and one sitting exactly at the cutoff boundary, from either side
    const justInside = join(dir, orphanName('other.pdf'))
    await writeFile(justInside, 'x')
    await utimes(
      justInside,
      new Date(now - SAVE_TEMP_TTL_MS + 60_000),
      new Date(now - SAVE_TEMP_TTL_MS + 60_000),
    )
    const justOutside = join(dir, orphanName('third.pdf'))
    await writeFile(justOutside, 'x')
    await utimes(
      justOutside,
      new Date(now - SAVE_TEMP_TTL_MS - 1),
      new Date(now - SAVE_TEMP_TTL_MS - 1),
    )

    const removed = await sweepStaleSaveTemps(dir, now)

    expect(removed).toEqual([justOutside])
    expect(existsSync(live)).toBe(true)
    expect(existsSync(justInside)).toBe(true)
  })

  it('touches only the save temp signature: targets, plain .tmp and off-pattern names stay', async () => {
    const dir = await tempDir()
    const now = Date.now()
    const stale = new Date(now - SAVE_TEMP_TTL_MS - 60_000)
    const arbitrary = [
      'report.pdf', // the picked target itself
      'report.pdf.tmp', // not dot-prefixed
      '.report.pdf.abc123.tmp', // 6 hex, not 12
      '.report.pdf.ab12cd34ef56e.tmp', // 13 chars
      '.report.pdf.AB12CD34EF56.tmp', // uppercase hex
      '.report.pdf.ab12cd34ef56.tmp.bak', // wrong suffix
    ]
    for (const name of arbitrary) {
      const path = join(dir, name)
      await writeFile(path, 'keep me')
      await utimes(path, stale, stale)
    }
    // a directory named exactly like an orphan must not be touched
    const dirLike = join(dir, orphanName('folder.pdf'))
    await mkdir(dirLike, { recursive: true })
    await writeFile(join(dirLike, 'inner.txt'), 'notes')

    const removed = await sweepStaleSaveTemps(dir, now)

    expect(removed).toEqual([])
    for (const name of arbitrary) expect(existsSync(join(dir, name))).toBe(true)
    expect(existsSync(join(dirLike, 'inner.txt'))).toBe(true)
  })

  it('sweeps nothing and does not throw when the directory is missing', async () => {
    const dir = await tempDir()
    const missing = join(dir, 'never-created')
    await expect(sweepStaleSaveTemps(missing)).resolves.toEqual([])
  })

  it('sweeps by directory, not by target name: any expired save temp in the dir goes', async () => {
    const dir = await tempDir()
    const now = Date.now()
    const stale = new Date(now - SAVE_TEMP_TTL_MS - 60_000)
    // a leftover of a save abandoned under a different target name
    const leftover = join(dir, orphanName('old-name.pdf'))
    await writeFile(leftover, 'orphaned bytes')
    await utimes(leftover, stale, stale)

    // the new save targets a different file in the same directory
    expect(await sweepStaleSaveTemps(dir, now)).toEqual([leftover])
    expect(existsSync(leftover)).toBe(false)
  })
})

describe('the sweep pattern matches the atomicWriteFile temp shape', () => {
  it('a temp built the production way (randomBytes(6) hex) is picked up by the sweep once stale', async () => {
    const dir = await tempDir()
    // exactly how atomic-write.ts names its temps: dot-prefixed, 12 lowercase hex
    const tmp = join(dir, `.doc.pdf.${randomBytes(6).toString('hex')}.tmp`)
    await writeFile(tmp, new Uint8Array([1, 2, 3]))
    const now = Date.now()
    // brand-new temp survives (the save below it is live)…
    expect(await sweepStaleSaveTemps(dir, now)).toEqual([])
    expect(existsSync(tmp)).toBe(true)
    // …then ages past the TTL (kill -9) and is swept at the next save
    const stale = new Date(now - SAVE_TEMP_TTL_MS - 60_000)
    await utimes(tmp, stale, stale)
    expect(await sweepStaleSaveTemps(dir, now + 1)).toEqual([tmp])
    expect(existsSync(tmp)).toBe(false)
  })
})

describe('savePdfToPath sweeps expired temps (OBS-1670)', () => {
  it('the next save removes a kill -9 orphan next to the target and still writes cleanly', async () => {
    const dir = await tempDir()
    const src = join(dir, 'original.pdf')
    const dst = join(dir, 'report.pdf')
    await writeFile(src, await makePdf())
    const leftover = join(dir, orphanName('report.pdf'))
    await writeFile(leftover, 'orphaned bytes')
    const now = Date.now()
    const stale = new Date(now - SAVE_TEMP_TTL_MS - 60_000)
    await utimes(leftover, stale, stale)

    await savePdfToPath(src, dst, request(src, { targetPath: dst }))

    expect(existsSync(leftover)).toBe(false)
    // the save itself committed: a valid PDF, no temps left behind
    const out = await PDFDocument.load(new Uint8Array(readFileSync(dst)))
    expect(out.getPageCount()).toBe(1)
    expect(basename(dst)).toBe('report.pdf')
  })

  it('a fresh in-flight temp in the target directory survives the save', async () => {
    const dir = await tempDir()
    const src = join(dir, 'original.pdf')
    const dst = join(dir, 'report.pdf')
    await writeFile(src, await makePdf())
    // a concurrent save's temp written seconds ago must not be a candidate
    const live = join(dir, orphanName('report.pdf'))
    await writeFile(live, 'writing…')
    const fresh = new Date(Date.now() - 5_000)
    await utimes(live, fresh, fresh)

    await savePdfToPath(src, dst, request(src, { targetPath: dst }))

    expect(existsSync(live)).toBe(true)
    expect(existsSync(dst)).toBe(true)
  })
})

describe('save-pdf.ts wiring (source contract)', () => {
  const here = dirname(fileURLToPath(import.meta.url))
  const source = readFileSync(join(here, '../src/main/save-pdf.ts'), 'utf8')

  it('every fresh save sweeps the target directory before atomicWriteFile creates its temp', () => {
    const sweepAt = source.indexOf('await sweepStaleSaveTemps(dirname(targetPath))')
    const writeAt = source.indexOf('await atomicWriteFile(targetPath, bytes)')
    expect(sweepAt).toBeGreaterThan(-1)
    expect(writeAt).toBeGreaterThan(sweepAt)
  })
})

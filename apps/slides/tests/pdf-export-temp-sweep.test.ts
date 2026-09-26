/**
 * BUG-1767 (audit EXP-4): a kill -9 mid PDF-export bypasses exportSlidesPdf's
 * finally-block, so the mkdtemp print-HTML directory (multi-MB slides.html on
 * large decks) survives in the shared OS temp root forever — nothing revisits
 * this prefix at startup, unlike the video streaming temps (OBS-1658) and the
 * generated-page dirs. The fix: each fresh PDF export (slides:export-pdf)
 * sweeps the expired airy-slides-pdf-* directories first, age alone deciding
 * so a live export is never a candidate.
 *
 * The sweep lives in pdf-export-temp.ts (importable, unlike slides-main) and
 * the wiring is pinned source-style, as in video-export-temp-sweep.test.ts.
 */
import { existsSync, readFileSync, symlinkSync, writeFileSync } from 'node:fs'
import { mkdir, mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'

import {
  PDF_EXPORT_TEMP_PREFIX,
  PDF_EXPORT_TEMP_TTL_MS,
  sweepStalePdfExportTempDirs,
} from '../src/main/pdf-export-temp'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function tempRoot(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'airy-slides-pdf-sweep-test-'))
  roots.push(dir)
  return dir
}

/** An orphan dir of the exact shape exportSlidesPdf's mkdtemp produces. */
const orphanDir = (root: string, suffix = 'Ab3xY9') =>
  join(root, `${PDF_EXPORT_TEMP_PREFIX}${suffix}`)

/** Populate a temp dir the way a killed export leaves it. */
async function makeOrphan(path: string, htmlBytes = 'deck pages…'): Promise<void> {
  await mkdir(path, { recursive: true })
  await writeFile(join(path, 'slides.html'), htmlBytes, 'utf8')
}

describe('stale PDF-export temp dir sweep (BUG-1767)', () => {
  it('removes an expired kill -9 leftover directory with its contents', async () => {
    const root = await tempRoot()
    const leftover = orphanDir(root)
    await makeOrphan(leftover, '<html>'.padEnd(4096, 'x'))
    const now = Date.now()
    const stale = new Date(now - PDF_EXPORT_TEMP_TTL_MS - 60_000)
    await utimes(leftover, stale, stale)

    const removed = await sweepStalePdfExportTempDirs(root, now)

    expect(removed).toEqual([leftover])
    expect(existsSync(leftover)).toBe(false)
  })

  it('never sweeps a live export dir: anything within the TTL window stays', async () => {
    const root = await tempRoot()
    const now = Date.now()
    // a dir written seconds ago (in-progress export, printToPDF still running)
    const live = orphanDir(root, 'Live01')
    await makeOrphan(live)
    const fresh = new Date(now - 5_000)
    await utimes(live, fresh, fresh)
    // one sitting exactly at the cutoff boundary, from either side
    const justInside = orphanDir(root, 'Inside')
    await makeOrphan(justInside)
    await utimes(
      justInside,
      new Date(now - PDF_EXPORT_TEMP_TTL_MS + 60_000),
      new Date(now - PDF_EXPORT_TEMP_TTL_MS + 60_000),
    )
    const justOutside = orphanDir(root, 'Border')
    await makeOrphan(justOutside)
    await utimes(
      justOutside,
      new Date(now - PDF_EXPORT_TEMP_TTL_MS - 1),
      new Date(now - PDF_EXPORT_TEMP_TTL_MS - 1),
    )

    const removed = await sweepStalePdfExportTempDirs(root, now)

    expect(removed).toEqual([justOutside])
    expect(existsSync(live)).toBe(true)
    expect(existsSync(justInside)).toBe(true)
  })

  it('touches only the suite signature: look-alike names, files and nested paths stay', async () => {
    const root = await tempRoot()
    const now = Date.now()
    const stale = new Date(now - PDF_EXPORT_TEMP_TTL_MS - 60_000)
    const arbitrary = [
      'airy-slides-pdf-', // no mkdtemp suffix at all
      'airy-slides-pdf-abc12', // 5 chars, not 6
      'airy-slides-pdf-abc1234', // 7 chars
      'airy-slides-pdf-abc12-', // separator inside the suffix
      'airy-slides-pdf-abc123.tmp', // wrong suffix
      'airy-slides-pdf-other', // plain look-alike with a dash
    ]
    for (const name of arbitrary) {
      const path = join(root, name)
      await mkdir(path, { recursive: true })
      await writeFile(join(path, 'keep.txt'), 'keep me')
      await utimes(path, stale, stale)
    }
    // a plain FILE named exactly like an orphan must not be touched
    const fileLike = join(root, `${PDF_EXPORT_TEMP_PREFIX}File01`)
    writeFileSync(fileLike, 'a file, not a dir')
    await utimes(fileLike, stale, stale)
    // a symlink entry named like an orphan must not be followed or removed:
    // its target is a plain (non-matching) directory nothing else would sweep
    const target = join(root, 'sweep-test-target')
    await makeOrphan(target)
    await utimes(target, stale, stale)
    const link = join(root, `${PDF_EXPORT_TEMP_PREFIX}Link01`)
    symlinkSync(target, link, 'dir')
    // a matching orphan nested one level deeper is out of the sweep's reach
    const nested = join(root, 'sub', `${PDF_EXPORT_TEMP_PREFIX}abc123`)
    await makeOrphan(nested)
    await utimes(nested, stale, stale)

    const removed = await sweepStalePdfExportTempDirs(root, now)

    expect(removed).toEqual([])
    for (const name of arbitrary) expect(existsSync(join(root, name))).toBe(true)
    expect(existsSync(fileLike)).toBe(true)
    expect(existsSync(link)).toBe(true)
    expect(existsSync(join(target, 'slides.html'))).toBe(true)
    expect(existsSync(join(nested, 'slides.html'))).toBe(true)
  })

  it('sweeps every expired orphan of the prefix in one pass', async () => {
    const root = await tempRoot()
    const now = Date.now()
    const stale = new Date(now - PDF_EXPORT_TEMP_TTL_MS - 60_000)
    const orphans = [
      orphanDir(root, 'One001'),
      orphanDir(root, 'Two002'),
      orphanDir(root, 'Zz9zZz'),
    ]
    for (const orphan of orphans) {
      await makeOrphan(orphan)
      await utimes(orphan, stale, stale)
    }

    const removed = await sweepStalePdfExportTempDirs(root, now)

    expect(removed.sort()).toEqual(orphans.sort())
    for (const orphan of orphans) expect(existsSync(orphan)).toBe(false)
  })

  it('sweeps nothing and does not throw when the temp root is missing', async () => {
    const root = await tempRoot()
    const missing = join(root, 'never-created')
    await expect(sweepStalePdfExportTempDirs(missing)).resolves.toEqual([])
  })

  it('a real mkdtemp dir of the production prefix is picked up once stale', async () => {
    const root = await tempRoot()
    // built exactly the way exportSlidesPdf does it — pins the sweep regex to
    // mkdtemp's six-char [a-zA-Z0-9] suffix so the two cannot drift apart
    const dir = await mkdtemp(join(root, PDF_EXPORT_TEMP_PREFIX))
    await writeFile(join(dir, 'slides.html'), 'pages', 'utf8')
    const now = Date.now()
    // brand-new dir survives (the export below it is live)…
    expect(await sweepStalePdfExportTempDirs(root, now)).toEqual([])
    expect(existsSync(dir)).toBe(true)
    // …then ages past the TTL (kill -9) and is swept at the next export
    const stale = new Date(now - PDF_EXPORT_TEMP_TTL_MS - 60_000)
    await utimes(dir, stale, stale)
    expect(await sweepStalePdfExportTempDirs(root, now + 1)).toEqual([dir])
    expect(existsSync(dir)).toBe(false)
  })

  it('the TTL is the one hour the audit prescribes', () => {
    expect(PDF_EXPORT_TEMP_TTL_MS).toBe(60 * 60 * 1000)
  })
})

describe('sweep wiring (source contract, slides-main is electron-only)', () => {
  const here = dirname(fileURLToPath(import.meta.url))
  const mainSource = readFileSync(join(here, '../src/main/slides-main.ts'), 'utf8')
  const pdfSource = readFileSync(join(here, '../src/main/pdf-export.ts'), 'utf8')

  it('every fresh PDF export sweeps the OS temp root before mkdtemp', () => {
    expect(mainSource).toContain('void sweepStalePdfExportTempDirs()')
  })

  it('the temp dir is built from the shared prefix owner, not an inline literal', () => {
    expect(pdfSource).toContain('mkdtemp(join(tmpdir(), PDF_EXPORT_TEMP_PREFIX))')
    // the inline literal the sweep pattern mirrors is gone from pdf-export
    expect(pdfSource).not.toContain("'airy-slides-pdf-'")
  })
})

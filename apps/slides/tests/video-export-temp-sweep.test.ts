/**
 * OBS-1658: a kill -9 mid video-export bypasses every cleanup hook
 * (render-process-gone / destroyed never fire on SIGKILL), so the
 * dot-prefixed `.target.<12 hex>.tmp` streaming temp survives in the
 * user-picked directory — a directory nothing app-owned ever revisits. The
 * fix: each fresh export into a directory (video-file-stream-begin) sweeps
 * that directory's expired video temps first. Age alone decides, so a live
 * export's continuously appended temp is never a candidate.
 *
 * The sweep lives in video-export-temp.ts (importable, unlike slides-main)
 * and the begin-handler wiring is pinned source-style, as in
 * video-export-throttle-restore.test.ts.
 */
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'

import {
  sweepStaleVideoExportTemps,
  VIDEO_EXPORT_TEMP_TTL_MS,
  videoExportTempPath,
} from '../src/main/video-export-temp'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function tempDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'airy-slides-video-tmp-test-'))
  roots.push(dir)
  return dir
}

/** A temp name of the exact shape slides-main streamed exports produce. */
const orphanName = (target: string) => `.${target}.${'ab12cd34ef56'}.tmp`

describe('stale video-export temp sweep (OBS-1658)', () => {
  it('removes an expired kill -9 leftover temp next to the export target', async () => {
    const dir = await tempDir()
    const leftover = join(dir, orphanName('holiday.mp4'))
    await writeFile(leftover, 'orphaned bytes')
    const now = Date.now()
    const stale = new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 60_000)
    await utimes(leftover, stale, stale)

    const removed = await sweepStaleVideoExportTemps(dir, now)

    expect(removed).toEqual([leftover])
    expect(existsSync(leftover)).toBe(false)
  })

  it('never sweeps a live export temp: anything within the TTL window stays', async () => {
    const dir = await tempDir()
    const now = Date.now()
    // a temp written seconds ago (in-progress export, mtime moves per append)
    const live = join(dir, orphanName('holiday.mp4'))
    await writeFile(live, 'streaming…')
    const fresh = new Date(now - 5_000)
    await utimes(live, fresh, fresh)
    // and one sitting exactly at the cutoff boundary, from either side
    const justInside = join(dir, orphanName('other.webm'))
    await writeFile(justInside, 'x')
    await utimes(
      justInside,
      new Date(now - VIDEO_EXPORT_TEMP_TTL_MS + 60_000),
      new Date(now - VIDEO_EXPORT_TEMP_TTL_MS + 60_000),
    )
    const justOutside = join(dir, orphanName('third.webm'))
    await writeFile(justOutside, 'x')
    await utimes(
      justOutside,
      new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 1),
      new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 1),
    )

    const removed = await sweepStaleVideoExportTemps(dir, now)

    expect(removed).toEqual([justOutside])
    expect(existsSync(live)).toBe(true)
    expect(existsSync(justInside)).toBe(true)
  })

  it('touches only the suite temp signature: targets, plain .tmp and off-pattern names stay', async () => {
    const dir = await tempDir()
    const now = Date.now()
    const stale = new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 60_000)
    const arbitrary = [
      'holiday.mp4', // the picked target itself
      'holiday.mp4.tmp', // not dot-prefixed
      '.holiday.mp4.abc123.tmp', // 6 hex, not 12
      '.holiday.mp4.ab12cd34ef56e.tmp', // 13 chars
      '.holiday.mp4.AB12CD34EF56.tmp', // uppercase hex
      '.holiday.mp4.ab12cd34ef56.tmp.bak', // wrong suffix
    ]
    for (const name of arbitrary) {
      const path = join(dir, name)
      await writeFile(path, 'keep me')
      await utimes(path, stale, stale)
    }
    // a directory named exactly like an orphan must not be touched
    const dirLike = join(dir, orphanName('folder.mp4'))
    await mkdir(dirLike, { recursive: true })
    await writeFile(join(dirLike, 'inner.txt'), 'notes')

    const removed = await sweepStaleVideoExportTemps(dir, now)

    expect(removed).toEqual([])
    for (const name of arbitrary) expect(existsSync(join(dir, name))).toBe(true)
    expect(existsSync(join(dirLike, 'inner.txt'))).toBe(true)
  })

  it('sweeps nothing and does not throw when the directory is missing', async () => {
    const dir = await tempDir()
    const missing = join(dir, 'never-created')
    await expect(sweepStaleVideoExportTemps(missing)).resolves.toEqual([])
  })

  it('sweeps by directory, not by target name: any expired suite temp in the dir goes', async () => {
    const dir = await tempDir()
    const now = Date.now()
    const stale = new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 60_000)
    // a leftover of an export abandoned under a different target name
    const leftover = join(dir, orphanName('old-name.webm'))
    await writeFile(leftover, 'orphaned bytes')
    await utimes(leftover, stale, stale)

    // the new export targets a different file in the same directory
    expect(await sweepStaleVideoExportTemps(dir, now)).toEqual([leftover])
    expect(existsSync(leftover)).toBe(false)
  })
})

describe('videoExportTempPath matches the sweep pattern', () => {
  it('produces dot-prefixed `<target>.<12 hex>.tmp` in the target directory', async () => {
    const dir = await tempDir()
    const target = join(dir, 'my.movie.final.mp4')
    const tmp = videoExportTempPath(target)
    expect(dirname(tmp)).toBe(dir)
    expect(basename(tmp)).toMatch(/^\..+\.[0-9a-f]{12}\.tmp$/)
    // the temp is hidden next to the target and unique per export
    expect(tmp.startsWith(join(dir, '.'))).toBe(true)
    expect(videoExportTempPath(target)).not.toBe(tmp)
  })

  it('a real temp created the production way is picked up by the sweep once stale', async () => {
    const dir = await tempDir()
    const tmp = videoExportTempPath(join(dir, 'deck.mp4'))
    await writeFile(tmp, new Uint8Array([1, 2, 3]))
    const now = Date.now()
    // brand-new temp survives (the export below it is live)…
    expect(await sweepStaleVideoExportTemps(dir, now)).toEqual([])
    expect(existsSync(tmp)).toBe(true)
    // …then ages past the TTL (kill -9) and is swept at the next export
    const stale = new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 60_000)
    await utimes(tmp, stale, stale)
    expect(await sweepStaleVideoExportTemps(dir, now + 1)).toEqual([tmp])
    expect(existsSync(tmp)).toBe(false)
  })
})

describe('video-file-stream-begin wiring (source contract, slides-main is electron-only)', () => {
  const here = dirname(fileURLToPath(import.meta.url))
  const source = readFileSync(join(here, '../src/main/slides-main.ts'), 'utf8')

  it('every fresh export sweeps the target directory before opening its temp', () => {
    expect(source).toContain('void sweepStaleVideoExportTemps(dirname(filePath))')
  })

  it('the streaming temp is built by the shared pattern owner, not inline', () => {
    expect(source).toContain('const tmp = videoExportTempPath(filePath)')
    // the inline randomBytes construction the sweep pattern mirrors is gone
    expect(source).not.toContain("randomBytes(6).toString('hex')")
  })
})

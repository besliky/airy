/**
 * BUG-1767 (audit EXP-3): the image-series export wrote each member with a
 * direct writeFile, so a kill -9 mid-series left a partial series with a
 * TRUNCATED member (9 valid + 1 undecodable) and no marker of any kind. The
 * fix commits each member through the shared atomicWriteFile (same-dir temp +
 * rename): already-written members stay valid, the interrupted member is
 * simply absent, and its temp falls off as a dot-prefixed orphan that the
 * next export into the directory sweeps (the OBS-1658 pattern — the temp
 * shape is shared with atomicWriteFile, so the video sweep covers it).
 *
 * The handler wiring is pinned source-style (slides-main is electron-only,
 * as in video-export-temp-sweep.test.ts); the kill itself is simulated at
 * the exact state SIGKILL during a member write leaves behind.
 */
import { randomBytes } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { mkdtemp, readdir, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'

import { atomicWriteFile } from '@airy-office/electron-utils'
import { sweepStaleVideoExportTemps, VIDEO_EXPORT_TEMP_TTL_MS } from '../src/main/video-export-temp'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function exportDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'airy-slides-images-test-'))
  roots.push(dir)
  return dir
}

/** A synthetic JPEG-sized member: SOI … payload … EOI, unique per slide. */
function jpegBytes(slide: number, size = 64 * 1024): Uint8Array {
  const bytes = new Uint8Array(size)
  bytes[0] = 0xff
  bytes[1] = 0xd8 // SOI
  bytes[2] = 0xff
  bytes[3] = 0xe0 // APP0
  for (let i = 4; i < size - 2; i++) bytes[i] = (slide * 31 + i) % 256
  bytes[size - 2] = 0xff
  bytes[size - 1] = 0xd9 // EOI
  return bytes
}

/** The decoder-side check the audit ran with PIL: intact frame, exact bytes. */
function expectValidJpeg(path: string, expected: Uint8Array): void {
  const bytes = readFileSync(path)
  expect(bytes[0]).toBe(0xff)
  expect(bytes[1]).toBe(0xd8)
  expect(bytes[bytes.length - 2]).toBe(0xff)
  expect(bytes[bytes.length - 1]).toBe(0xd9)
  expect(bytes.equals(expected)).toBe(true)
}

const memberPath = (dir: string, n: number) => join(dir, `deck-${String(n).padStart(3, '0')}.jpg`)

/** A temp of the exact shape atomicWriteFile creates next to a member. */
const orphanTempName = (member: string) => `.${member}.${randomBytes(6).toString('hex')}.tmp`

describe('kill -9 mid JPEG-series: atomicWriteFile members (BUG-1767)', () => {
  it('a kill during member N leaves members 1..N-1 valid, member N absent', async () => {
    const dir = await exportDir()
    const bytes = new Map<number, Uint8Array>()
    // the series up to the kill, committed the production way
    for (let n = 1; n <= 9; n++) {
      const data = jpegBytes(n)
      bytes.set(n, data)
      await atomicWriteFile(memberPath(dir, n), data)
    }
    // SIGKILL lands mid-write of member 10: a partial temp with a truncated
    // tail exists next to the target, the rename never happens
    const interrupted = jpegBytes(10)
    bytes.set(10, interrupted)
    const truncated = interrupted.subarray(0, Math.floor(interrupted.length / 2))
    const member10 = memberPath(dir, 10)
    const orphan = join(dir, orphanTempName(basename(member10)))
    await writeFile(orphan, truncated)

    // the already-written members decode intact (the audit's "9 valid")
    for (let n = 1; n <= 9; n++) expectValidJpeg(memberPath(dir, n), bytes.get(n)!)
    // the interrupted member never surfaces as a broken file
    expect(existsSync(member10)).toBe(false)
    // the interrupted write fell off as a temp, not as a member
    expect((await readdir(dir)).filter((name) => name.endsWith('.tmp'))).toEqual([basename(orphan)])
    expect(readFileSync(orphan).length).toBe(truncated.length)
  })

  it('a completed series leaves zero temps: every member committed, none partial', async () => {
    const dir = await exportDir()
    for (let n = 1; n <= 4; n++) await atomicWriteFile(memberPath(dir, n), jpegBytes(n))

    expect((await readdir(dir)).filter((name) => name.endsWith('.tmp'))).toEqual([])
    for (let n = 1; n <= 4; n++) expectValidJpeg(memberPath(dir, n), jpegBytes(n))
  })

  it('the next export sweeps the interrupted temp: members and fresh temps stay', async () => {
    const dir = await exportDir()
    const member1 = memberPath(dir, 1)
    await atomicWriteFile(member1, jpegBytes(1))
    // the orphan a kill -9 left behind, now aged past the sweep TTL
    const orphan = join(dir, orphanTempName('deck-002.jpg'))
    await writeFile(orphan, 'truncated tail')
    const now = Date.now()
    const stale = new Date(now - VIDEO_EXPORT_TEMP_TTL_MS - 60_000)
    await utimes(orphan, stale, stale)
    // a concurrent live series' temp is seconds old — never a candidate
    const liveTemp = join(dir, orphanTempName('deck-003.jpg'))
    await writeFile(liveTemp, 'streaming member')

    const removed = await sweepStaleVideoExportTemps(dir, now)

    expect(removed).toEqual([orphan])
    expect(existsSync(orphan)).toBe(false)
    expect(existsSync(member1)).toBe(true)
    expect(existsSync(liveTemp)).toBe(true)
  })
})

describe('export-images wiring (source contract, slides-main is electron-only)', () => {
  const here = dirname(fileURLToPath(import.meta.url))
  const source = readFileSync(join(here, '../src/main/slides-main.ts'), 'utf8')

  it('every series member commits through atomicWriteFile, not a direct writeFile', () => {
    expect(source).toContain('await atomicWriteFile(paths[i]')
    // the non-atomic member write the audit caught is gone
    expect(source).not.toContain('await writeFile(paths[i]')
  })

  it('the fresh export sweeps the picked directory before writing members', () => {
    expect(source).toContain('void sweepStaleVideoExportTemps(op.dir)')
  })
})

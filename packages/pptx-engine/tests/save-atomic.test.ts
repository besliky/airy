import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import type { Readable } from 'node:stream'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type JSZip from 'jszip'
import { openPptx, savePptxToFile } from '../src/index'

/**
 * Atomicity of savePptxToFile (BUG-402): the streamed zip must land in a
 * same-directory `.<name>.<rand>.tmp` file promoted over the target by rename,
 * so a crash or ENOSPC mid-write can never truncate the user's deck.
 *
 * Failure injection follows packages/electron-utils/tests/atomic-write.test.ts:
 * node:fs/promises.rename is wrapped for promote failures, and jszip's node
 * stream is wrapped in a Readable that errors mid-stream — after forwarding the
 * leading chunks, so the pipeline really dies with partial bytes on disk.
 */
let mockFailSourceAfterBytes = -1

vi.mock('jszip', async (importOriginal) => {
  const { Readable: ReadableStream } = await import('node:stream')
  const actual = (await importOriginal()) as unknown as { default: typeof JSZip }
  const RealJSZip = actual.default
  return {
    default: class extends RealJSZip {
      generateNodeStream(...args: Parameters<JSZip['generateNodeStream']>): NodeJS.ReadableStream {
        const real = super.generateNodeStream(...args) as Readable
        if (mockFailSourceAfterBytes < 0) return real
        const limit = mockFailSourceAfterBytes
        mockFailSourceAfterBytes = -1
        const failing = new ReadableStream({ read() {} })
        let seen = 0
        real.on('data', (chunk: Buffer) => {
          if (seen >= limit) return
          seen += chunk.length
          failing.push(chunk)
          if (seen >= limit) {
            // Give the pushed chunk a tick to flow into the destination, so the
            // temp file really holds partial bytes when the pipeline dies.
            setImmediate(() => {
              real.destroy()
              failing.destroy(new Error('simulated source failure mid-stream'))
            })
          }
        })
        real.on('end', () => failing.push(null))
        real.on('error', (error: Error) => failing.destroy(error))
        return failing
      }
    },
  }
})

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename) }
})

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))
const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04])

describe('savePptxToFile atomicity', () => {
  let dir = ''

  afterEach(() => {
    if (dir) rmSync(dir, { recursive: true, force: true })
    dir = ''
    mockFailSourceAfterBytes = -1
    vi.mocked(rename).mockReset()
  })

  it('writes a valid package and leaves no temp file behind', async () => {
    dir = mkdtempSync(join(tmpdir(), 'save-atomic-'))
    const target = join(dir, 'deck.pptx')
    const opened = await openPptx(fx('01_standard_business.pptx'))

    await savePptxToFile(opened, target)

    const reopened = await openPptx(readFileSync(target))
    expect(reopened.deck.slides.length).toBe(opened.deck.slides.length)
    expect(readdirSync(dir)).toEqual(['deck.pptx'])
  })

  it('keeps the target byte-identical and cleans the temp when the source dies mid-stream', async () => {
    dir = mkdtempSync(join(tmpdir(), 'save-atomic-'))
    const target = join(dir, 'deck.pptx')
    const before = 'previous deck bytes that must survive a failed save'
    writeFileSync(target, before)
    const opened = await openPptx(fx('01_standard_business.pptx'))

    mockFailSourceAfterBytes = 64
    await expect(savePptxToFile(opened, target)).rejects.toThrow(
      'simulated source failure mid-stream',
    )

    expect(readFileSync(target, 'utf-8')).toBe(before)
    expect(readdirSync(dir)).toEqual(['deck.pptx'])
  })

  it('replaces the contents of an existing target on success', async () => {
    dir = mkdtempSync(join(tmpdir(), 'save-atomic-'))
    const target = join(dir, 'deck.pptx')
    writeFileSync(target, 'stale bytes')
    const opened = await openPptx(fx('01_standard_business.pptx'))

    await savePptxToFile(opened, target)

    const bytes = readFileSync(target)
    expect(bytes.subarray(0, 4)).toEqual(ZIP_MAGIC)
    const reopened = await openPptx(bytes)
    expect(reopened.deck.slides.length).toBe(opened.deck.slides.length)
    expect(readdirSync(dir)).toEqual(['deck.pptx'])
  })

  it('keeps the target intact and cleans the temp when the promote rename fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'save-atomic-'))
    const target = join(dir, 'deck.pptx')
    const before = 'previous deck bytes that must survive a failed promote'
    writeFileSync(target, before)
    const opened = await openPptx(fx('01_standard_business.pptx'))

    vi.mocked(rename).mockRejectedValueOnce(
      Object.assign(new Error('EACCES: rename locked by another program'), { code: 'EACCES' }),
    )
    await expect(savePptxToFile(opened, target)).rejects.toThrow('EACCES')

    expect(readFileSync(target, 'utf-8')).toBe(before)
    expect(readdirSync(dir)).toEqual(['deck.pptx'])
  })
})

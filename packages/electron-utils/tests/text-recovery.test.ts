import { mkdtempSync, readFileSync, rmSync, statSync, utimesSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * TextRecoveryStore (src/text-recovery.ts): the crash-recovery copy scheme
 * shared by the markdown and html editors — path mapping, the newer-than
 * decision, save/discard cleanup, and the clear/write race that must not
 * resurrect a discarded copy.
 */
import {
  TextRecoveryStore,
  shouldOfferTextRecovery,
  textRecoveryPathFor,
} from '../src/text-recovery'

let scratch: string
let store: TextRecoveryStore

beforeEach(() => {
  scratch = mkdtempSync(join(tmpdir(), 'airy-text-recovery-'))
  store = new TextRecoveryStore(scratch, {
    write: async (path, text) => void writeFileSync(path, text, 'utf8'),
    // decode-free read: the store treats original and copy identically
    readOriginal: async (path) => readFileSync(path, 'utf8'),
  })
})

afterEach(() => {
  rmSync(scratch, { recursive: true, force: true })
})

describe('textRecoveryPathFor', () => {
  it('names copies by a digest of the document path, inside the store dir', () => {
    const copy = textRecoveryPathFor('/docs', '/docs/notes.md')
    expect(copy.startsWith(join('/docs'))).toBe(true)
    expect(copy.endsWith('.recovery')).toBe(true)
    // user paths never leak into the filename (path separators, unicode…)
    expect(copy).not.toContain('notes.md')
  })

  it('maps the same document path to the same copy path', () => {
    expect(textRecoveryPathFor('/d', '/a.md')).toBe(textRecoveryPathFor('/d', '/a.md'))
    expect(textRecoveryPathFor('/d', '/a.md')).not.toBe(textRecoveryPathFor('/d', '/b.md'))
  })
})

describe('shouldOfferTextRecovery', () => {
  it('offers only strictly newer copies', () => {
    expect(shouldOfferTextRecovery(2000, 1000)).toBe(true)
    expect(shouldOfferTextRecovery(1000, 1000)).toBe(false)
    expect(shouldOfferTextRecovery(500, 1000)).toBe(false)
  })
})

describe('writeCopy / clear / maybeRecover', () => {
  it('round-trips a copy through maybeRecover when the user restores', async () => {
    const file = join(scratch, 'doc.md')
    writeFileSync(file, 'saved', 'utf8')
    await store.writeCopy(file, 'recovered edits')
    // make the saved file older than the copy
    const past = new Date(Date.now() - 60_000)
    utimesSync(file, past, past)

    const prompt = vi.fn(async () => 'restore' as const)
    const result = await store.maybeRecover(file, prompt)
    expect(result).toEqual({ text: 'recovered edits', recovered: true })
    expect(prompt).toHaveBeenCalledTimes(1)
  })

  it('serves the original and deletes the copy when the user discards', async () => {
    const file = join(scratch, 'doc.md')
    writeFileSync(file, 'saved', 'utf8')
    await store.writeCopy(file, 'recovered edits')
    // make the saved file older than the copy
    const past = new Date(Date.now() - 60_000)
    utimesSync(file, past, past)

    const result = await store.maybeRecover(file, async () => 'discard')
    expect(result).toEqual({ text: 'saved', recovered: false })
    expect(() => statSync(store.pathFor(file))).toThrow() // gone
  })

  it('never prompts when the copy is older or absent', async () => {
    const file = join(scratch, 'doc.md')
    writeFileSync(file, 'saved', 'utf8')
    await store.writeCopy(file, 'stale edits')
    const past = new Date(Date.now() - 60_000)
    utimesSync(store.pathFor(file), past, past) // copy older than the freshly written file

    const prompt = vi.fn(async () => 'restore' as const)
    expect(await store.maybeRecover(file, prompt)).toEqual({ text: 'saved', recovered: false })
    expect(prompt).not.toHaveBeenCalled()

    store.clear(file)
    expect(await store.maybeRecover(file, prompt)).toEqual({ text: 'saved', recovered: false })
    expect(prompt).not.toHaveBeenCalled()
  })

  it('a clear racing an in-flight write does not resurrect the discarded copy', async () => {
    const file = join(scratch, 'doc.md')
    writeFileSync(file, 'saved', 'utf8')

    let releaseWrite!: () => void
    const gate = new Promise<void>((resolve) => (releaseWrite = resolve))
    const racing = new TextRecoveryStore(scratch, {
      write: async (path, text) => {
        await gate
        writeFileSync(path, text, 'utf8')
      },
      readOriginal: async (path) => originals.get(path) ?? 'saved',
    })
    const inFlight = racing.writeCopy(file, 'edits the user discarded')
    racing.clear(file) // discard lands while the write is pending
    releaseWrite()
    await inFlight
    // the copy must be gone: clear() won the race
    expect(() => statSync(racing.pathFor(file))).toThrow()
  })

  it('keeps the copy when the clear happened before the write started', async () => {
    const file = join(scratch, 'doc.md')
    writeFileSync(file, 'saved', 'utf8')
    store.clear(file)
    await store.writeCopy(file, 'fresh edits')
    expect(readFileSync(store.pathFor(file), 'utf8')).toBe('fresh edits')
  })
})

import { watch } from 'node:fs'
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  queueAppSettingsUpdate,
  readAppSettingsFile,
  writeAppSettingsFile,
} from '../src/app-settings-file'

/**
 * app-settings-file.ts — the single writer for userData/app-settings.json
 * (OBS-1532). Beyond the read/merge basics, these tests pin the concurrency
 * contract: concurrent read-modify-write cycles through the public API —
 * queued reducers, queued patches and synchronous merges interleaved with
 * random microtask/macrotask yields — must never drop or corrupt a key.
 */

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'app-settings-file-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('readAppSettingsFile', () => {
  it('returns an empty object when the file does not exist', () => {
    expect(readAppSettingsFile(settingsPath)).toEqual({})
  })

  it('returns an empty object for invalid JSON or a non-object root', () => {
    writeFileSync(settingsPath, 'not json')
    expect(readAppSettingsFile(settingsPath)).toEqual({})
    writeFileSync(settingsPath, '[1, 2]')
    expect(readAppSettingsFile(settingsPath)).toEqual({})
  })

  it('parses a valid settings object', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'zh', onboardingSeen: true }))
    expect(readAppSettingsFile(settingsPath)).toEqual({ language: 'zh', onboardingSeen: true })
  })
})

describe('writeAppSettingsFile (canonical synchronous merge)', () => {
  it('creates the file with the patch on first write', () => {
    writeAppSettingsFile(settingsPath, { onboardingSeen: true })
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ onboardingSeen: true })
  })

  it('merges into a fresh read, preserving unrelated keys', () => {
    writeAppSettingsFile(settingsPath, { language: 'ja' })
    writeAppSettingsFile(settingsPath, { onboardingSeen: true })
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'ja',
      onboardingSeen: true,
    })
  })

  it('recovers from a corrupt file by rewriting it', () => {
    writeFileSync(settingsPath, '{broken')
    writeAppSettingsFile(settingsPath, { language: 'en' })
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })

  it('leaves no temp files behind', () => {
    writeAppSettingsFile(settingsPath, { a: 1 })
    writeAppSettingsFile(settingsPath, { b: 2 })
    expect(readdirSync(dir).filter((name) => name.endsWith('.tmp'))).toEqual([])
  })
})

describe('queueAppSettingsUpdate (serialized async read-modify-write)', () => {
  it('applies a queued patch like the synchronous merge', async () => {
    writeAppSettingsFile(settingsPath, { language: 'de' })
    await queueAppSettingsUpdate(settingsPath, { onboardingSeen: true })
    expect(readAppSettingsFile(settingsPath)).toEqual({ language: 'de', onboardingSeen: true })
  })

  it('runs queued reducers strictly one at a time over fresh state', async () => {
    // The lost-update regression: 50 increments scheduled back-to-back must
    // all land. A wholesale write of a snapshot read before queueing, or two
    // overlapping sections, would lose some of them.
    const writes = Array.from({ length: 50 }, (_, i) =>
      queueAppSettingsUpdate(settingsPath, (current) => ({
        ...current,
        counter: typeof current.counter === 'number' ? current.counter + 1 : 1,
        last: i,
      })),
    )
    await Promise.all(writes)
    expect(readAppSettingsFile(settingsPath)).toEqual({ counter: 50, last: 49 })
  })

  it('lets synchronous merges interleave with the queue without losing keys', async () => {
    const queued = Array.from({ length: 30 }, (_, i) =>
      queueAppSettingsUpdate(settingsPath, { [`q${i}`]: i }),
    )
    for (let i = 0; i < 30; i++) writeAppSettingsFile(settingsPath, { [`s${i}`]: i })
    await Promise.all(queued)
    const stored = readAppSettingsFile(settingsPath)
    for (let i = 0; i < 30; i++) {
      expect(stored[`q${i}`]).toBe(i)
      expect(stored[`s${i}`]).toBe(i)
    }
  })

  it('rejects a failing update without poisoning the queue', async () => {
    const failing = queueAppSettingsUpdate(settingsPath, () => {
      throw new Error('boom')
    })
    await expect(failing).rejects.toThrow('boom')
    await queueAppSettingsUpdate(settingsPath, { recovered: true })
    expect(readAppSettingsFile(settingsPath)).toEqual({ recovered: true })
  })

  it('serializes independent queues per path', async () => {
    const otherPath = join(dir, 'other-settings.json')
    const a = queueAppSettingsUpdate(settingsPath, { a: 1 })
    const b = queueAppSettingsUpdate(otherPath, { b: 1 })
    await Promise.all([a, b])
    expect(readAppSettingsFile(settingsPath)).toEqual({ a: 1 })
    expect(readAppSettingsFile(otherPath)).toEqual({ b: 1 })
  })
})

// ── concurrency torture: concurrent RMW through the public API ──

/** deterministic PRNG (mulberry32) so failures replay exactly */
function mulberry32(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) >>> 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

describe('concurrent read-modify-write loses no keys (OBS-1532)', () => {
  // Hundreds of iterations per round: every write goes through the public
  // API with random key choices and random yields (microtasks, timers)
  // between schedules, mirroring the audit's interleaved probes. Keys are
  // never deleted, so the observed key set may only ever GROW — that
  // monotonic invariant is exactly what the transient key loss broke.
  const ROUNDS = 3
  const ITERATIONS = 250

  for (let round = 0; round < ROUNDS; round++) {
    it(`round ${round}: ${ITERATIONS} interleaved writes, monotonic key set`, async () => {
      const random = mulberry32(0x1532 + round)
      const observedKeys = new Set<string>()
      const writtenValues = new Map<string, Set<unknown>>()
      let pending: Array<Promise<void>> = []

      const record = (key: string, value: unknown) => {
        observedKeys.add(key)
        if (!writtenValues.has(key)) writtenValues.set(key, new Set())
        writtenValues.get(key)!.add(JSON.stringify(value))
      }

      const snapshot = () => {
        const stored = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
        // the atomic temp+rename must keep the file parseable at all times
        for (const key of observedKeys) expect(stored[key]).toBeDefined()
        return stored
      }

      writeAppSettingsFile(settingsPath, { starPrompt: { firstRunAt: 1 } })
      record('starPrompt', { firstRunAt: 1 })

      for (let i = 0; i < ITERATIONS; i++) {
        const key = `k${Math.floor(random() * 12)}`
        const value = `${round}:${i}`
        const kind = random()
        if (kind < 0.4) {
          // queued read-modify-write reducer (the dialog-memory shape)
          pending.push(
            queueAppSettingsUpdate(settingsPath, (current) => {
              record(key, value)
              return { ...current, [key]: value }
            }),
          )
        } else if (kind < 0.7) {
          // queued patch (the bridge toggle shape)
          record(key, value)
          pending.push(queueAppSettingsUpdate(settingsPath, { [key]: value }))
        } else {
          // synchronous merge (every shell settings handler shape)
          record(key, value)
          writeAppSettingsFile(settingsPath, { [key]: value })
        }
        // random interleaving: yield to microtasks/timers mid-stream
        if (random() < 0.1) {
          const drain = pending
          pending = []
          await Promise.resolve()
          if (random() < 0.5) await new Promise((resolve) => setTimeout(resolve, 0))
          await Promise.all(drain)
          snapshot()
        }
      }
      await Promise.all(pending)
      const stored = snapshot()
      // every key ever written is present, with one of its written values
      for (const [key, values] of writtenValues) {
        expect(values.has(JSON.stringify(stored[key]))).toBe(true)
      }
      expect(observedKeys.size).toBeGreaterThan(0)
    })
  }
})

// ── integration-style observer: what fs.watch sees on disk ──
// Watches the settings DIRECTORY (a file-level watch loses events when the
// writer's rename replaces the inode) and validates every observed state.
describe.skipIf(process.platform === 'win32')('fs.watch observer sees no key loss', () => {
  it('every on-disk state the watcher exposes is parseable and monotonic', async () => {
    const seenKeys = new Set<string>()
    let events = 0
    const watcher = watch(dir, (_event, filename) => {
      if (filename !== basename(settingsPath)) return
      const raw = readFileSync(settingsPath, 'utf8')
      const stored = JSON.parse(raw) as Record<string, unknown> // must always parse
      for (const key of seenKeys) expect(stored[key]).toBeDefined()
      for (const key of Object.keys(stored)) seenKeys.add(key)
      events += 1
    })
    try {
      const random = mulberry32(0x0b5e)
      let pending: Array<Promise<void>> = []
      for (let i = 0; i < 200; i++) {
        const key = `w${Math.floor(random() * 8)}`
        if (random() < 0.5) {
          pending.push(queueAppSettingsUpdate(settingsPath, { [key]: i }))
        } else {
          writeAppSettingsFile(settingsPath, { [key]: i })
        }
        if (random() < 0.15) {
          const drain = pending
          pending = []
          await new Promise((resolve) => setTimeout(resolve, 0))
          await Promise.all(drain)
        }
      }
      await Promise.all(pending)
      // give inotify a beat to deliver the tail events before closing
      await new Promise((resolve) => setTimeout(resolve, 50))
      expect(events).toBeGreaterThan(0)
      expect(seenKeys.size).toBeGreaterThan(0)
    } finally {
      watcher.close()
    }
  })
})

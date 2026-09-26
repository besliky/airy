import { watch } from 'node:fs'
import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  utimesSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  normalizeBooleanSetting,
  queueAppSettingsUpdate,
  readAppSettingsFile,
  recoverCorruptAppSettings,
  writeAppSettingsFile,
  type AppSettings,
} from '../src/app-settings-file'

/**
 * app-settings-file.ts — the single writer for userData/app-settings.json
 * (OBS-1532). Beyond the read/merge basics, these tests pin the concurrency
 * contract: concurrent read-modify-write cycles through the public API —
 * queued reducers, queued patches and synchronous merges interleaved with
 * random microtask/macrotask yields — must never drop or corrupt a key.
 * BUG-1771 adds the corrupt-file contract: corrupt bytes are preserved as
 * `.bak` before any rewrite and salvaged as far as strict JSON parsing
 * allows.
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

// ── BUG-1773: typed junk in known keys must not flip the decision ──
// The audited trap (SET-26-3): `restoreSession: "no"` counted as "restore on"
// under `!== false`. Every schema-boolean spelling resolves through
// normalizeBooleanSetting; junk falls back to the key's default.

describe('normalizeBooleanSetting (BUG-1773 typed-junk matrix)', () => {
  it('passes real booleans through', () => {
    expect(normalizeBooleanSetting(true, false)).toBe(true)
    expect(normalizeBooleanSetting(false, true)).toBe(false)
  })

  it('reads the false spellings a sloppy external editor writes', () => {
    for (const value of ['no', 'No', 'NO', 'false', 'FALSE', 'False', 'off', '0', '', '   ']) {
      expect(normalizeBooleanSetting(value, true), JSON.stringify(value)).toBe(false)
    }
  })

  it('reads the true spellings the same way', () => {
    for (const value of ['yes', 'Yes', 'true', 'TRUE', 'on', 'ON', '1']) {
      expect(normalizeBooleanSetting(value, false), JSON.stringify(value)).toBe(true)
    }
  })

  it('falls back to the schema default for typed junk and absent keys', () => {
    for (const value of ['banana', 'nope', 'enabled', 42, 0, null, undefined, {}, [], NaN]) {
      expect(normalizeBooleanSetting(value, true), JSON.stringify(value ?? 'null')).toBe(true)
      expect(normalizeBooleanSetting(value, false), JSON.stringify(value ?? 'null')).toBe(false)
    }
  })

  it('restoreSession semantics: the audited "no" string disables, absent stays on', () => {
    // mirrors the shell call shape (read + normalize with default on)
    writeFileSync(settingsPath, JSON.stringify({ restoreSession: 'no' }))
    expect(normalizeBooleanSetting(readAppSettingsFile(settingsPath).restoreSession, true)).toBe(
      false,
    )
    writeFileSync(settingsPath, JSON.stringify({ restoreSession: true }))
    expect(normalizeBooleanSetting(readAppSettingsFile(settingsPath).restoreSession, true)).toBe(
      true,
    )
    writeFileSync(settingsPath, JSON.stringify({ theme: 'dark' }))
    expect(normalizeBooleanSetting(readAppSettingsFile(settingsPath).restoreSession, true)).toBe(
      true,
    )
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

// ── BUG-1771: a corrupt file is preserved before any rewrite and salvaged ──

const backupPath = () => `${settingsPath}.bak`

/** silence + capture the forensic warnings for one test */
function captureWarnings(): { warnings: string[]; restore: () => void } {
  const warnings: string[] = []
  const spy = vi.spyOn(console, 'warn').mockImplementation((...args: unknown[]) => {
    warnings.push(args.map(String).join(' '))
  })
  return { warnings, restore: () => spy.mockRestore() }
}

/**
 * The seven corruption variants from the SET-26-1 audit. Repairable ones
 * must come back (fully or partially); wrong-type/empty roots have no object
 * content to salvage and honestly fall back to defaults. EVERY variant must
 * leave its exact corrupt bytes in the `.bak` — that copy is the only chance
 * a user has at forensic recovery once the merge-write replaces the file.
 */
const CORRUPT_VARIANTS: Array<{
  name: string
  content: string
  expected: AppSettings
  salvaged: boolean
}> = [
  {
    name: 'truncated JSON',
    content: '{"language":"de","theme":"dark","autoSaveDef',
    expected: { language: 'de', theme: 'dark' },
    salvaged: true,
  },
  {
    name: 'trailing garbage after a valid object',
    content: '{"language":"de","theme":"dark"} trailing garbage',
    expected: { language: 'de', theme: 'dark' },
    salvaged: true,
  },
  {
    name: 'UTF-8 BOM before a valid object',
    content: '\uFEFF{"language":"de","theme":"dark"}',
    expected: { language: 'de', theme: 'dark' },
    salvaged: true,
  },
  { name: 'empty file', content: '', expected: {}, salvaged: false },
  { name: 'JSON array root', content: '[1, 2, 3]', expected: {}, salvaged: false },
  { name: 'JSON number root', content: '42', expected: {}, salvaged: false },
  { name: 'whitespace-only file', content: '  \n\t ', expected: {}, salvaged: false },
]

describe('BUG-1771: corrupt app-settings.json forensics', () => {
  it('preserves every corrupt variant as .bak and salvages what parses', () => {
    for (const variant of CORRUPT_VARIANTS) {
      writeFileSync(settingsPath, variant.content)
      try {
        expect(readAppSettingsFile(settingsPath)).toEqual(variant.expected)
        const preserved = readFileSync(backupPath(), 'utf8')
        expect(preserved, `bak bytes for: ${variant.name}`).toBe(variant.content)
      } catch (error) {
        throw new Error(`variant failed: ${variant.name}`, { cause: error })
      } finally {
        rmSync(backupPath(), { force: true })
        rmSync(settingsPath, { force: true })
      }
    }
  })

  it('a valid file never gets a .bak and never logs', () => {
    const { warnings, restore } = captureWarnings()
    try {
      writeFileSync(settingsPath, JSON.stringify({ language: 'de' }))
      expect(readAppSettingsFile(settingsPath)).toEqual({ language: 'de' })
      expect(existsSync(backupPath())).toBe(false)
      expect(warnings).toEqual([])
    } finally {
      restore()
    }
  })

  it('a missing file gets no .bak (normal first run)', () => {
    expect(readAppSettingsFile(settingsPath)).toEqual({})
    expect(existsSync(backupPath())).toBe(false)
  })

  it('reading a corrupt file is side-effect free apart from the .bak', () => {
    // the corrupt file itself must stay byte-identical until a WRITE replaces
    // it — repair lands through the ordinary merge-write, never from a read
    const corrupt = CORRUPT_VARIANTS[0]!.content
    writeFileSync(settingsPath, corrupt)
    const { restore } = captureWarnings()
    try {
      readAppSettingsFile(settingsPath)
      expect(readFileSync(settingsPath, 'utf8')).toBe(corrupt)
    } finally {
      restore()
    }
  })

  it('logs the preservation once per distinct corrupt state, not per read', () => {
    writeFileSync(settingsPath, '{"language":"de"')
    const { warnings, restore } = captureWarnings()
    try {
      readAppSettingsFile(settingsPath)
      readAppSettingsFile(settingsPath)
      readAppSettingsFile(settingsPath)
      const preservedLogs = warnings.filter((line) => line.includes('preserved'))
      expect(preservedLogs).toHaveLength(1)
    } finally {
      restore()
    }
  })

  it('refreshes the .bak when the file corrupts again with different bytes', () => {
    const { restore } = captureWarnings()
    try {
      writeFileSync(settingsPath, '{a')
      readAppSettingsFile(settingsPath)
      // different SIZE guarantees a different fingerprint even in one ms
      writeFileSync(settingsPath, '{"much-longer-corrupt-payload":true,')
      readAppSettingsFile(settingsPath)
      expect(readFileSync(backupPath(), 'utf8')).toBe('{"much-longer-corrupt-payload":true,')
    } finally {
      restore()
    }
  })

  it('re-arms after the file becomes valid, preserving a later corruption too', () => {
    const { restore } = captureWarnings()
    try {
      writeFileSync(settingsPath, '{a')
      readAppSettingsFile(settingsPath)
      writeAppSettingsFile(settingsPath, { language: 'de' }) // file is healthy now
      readAppSettingsFile(settingsPath)
      expect(existsSync(backupPath())).toBe(true) // from the first corruption
      const firstBak = readFileSync(backupPath(), 'utf8')
      expect(firstBak).toBe('{a')
      writeFileSync(settingsPath, '[broken-beyond-repair')
      readAppSettingsFile(settingsPath)
      expect(readFileSync(backupPath(), 'utf8')).toBe('[broken-beyond-repair')
    } finally {
      restore()
    }
  })

  it('the .bak keeps the corrupt bytes own timestamps for forensics', () => {
    const yesterday = new Date(Date.now() - 86_400_000)
    writeFileSync(settingsPath, '{"language":"de"')
    utimesSync(settingsPath, yesterday, yesterday)
    const { restore } = captureWarnings()
    try {
      readAppSettingsFile(settingsPath)
      const preserved = statSync(backupPath())
      expect(Math.abs(preserved.mtimeMs - yesterday.getTime())).toBeLessThan(1000)
    } finally {
      restore()
    }
  })

  it('the merge-write persists salvaged keys — the audited total-loss repro', () => {
    // SET-26-1 live repro: seeded language + corrupt file + one theme toggle
    // used to end with ONLY the toggled key on disk; now language survives.
    writeFileSync(settingsPath, '{"language":"de","theme":"light","autoSaveDef')
    writeAppSettingsFile(settingsPath, { theme: 'dark' })
    expect(readAppSettingsFile(settingsPath)).toEqual({ language: 'de', theme: 'dark' })
  })

  it('the queued writer (#108) reads salvaged state and preserves the corrupt bytes', async () => {
    const corrupt = '{"language":"de","theme":"dark","autoSaveDef'
    writeFileSync(settingsPath, corrupt)
    await queueAppSettingsUpdate(settingsPath, (current) => ({
      ...current,
      onboardingSeen: true,
    }))
    expect(readAppSettingsFile(settingsPath)).toEqual({
      language: 'de',
      theme: 'dark',
      onboardingSeen: true,
    })
    expect(readFileSync(backupPath(), 'utf8')).toBe(corrupt)
    // the canonical write stays atomic: no temp leftovers
    expect(readdirSync(dir).filter((name) => name.endsWith('.tmp'))).toEqual([])
  })
})

describe('recoverCorruptAppSettings (strict-prefix salvage)', () => {
  const salvage = (text: string): AppSettings | null =>
    recoverCorruptAppSettings(text)?.settings ?? null

  it('completes a truncated tail keeping every finished member', () => {
    expect(salvage('{"language":"de","theme":"dark","autoSaveDef')).toEqual({
      language: 'de',
      theme: 'dark',
    })
    expect(salvage('{"a":1,"b":2,"c":3')).toEqual({ a: 1, b: 2, c: 3 })
    expect(salvage('{"a":1,')).toEqual({ a: 1 })
  })

  it('closes a string truncated mid-value', () => {
    expect(salvage('{"authorName":"Ada Lovelace","bio":"writ')).toEqual({
      authorName: 'Ada Lovelace',
      bio: 'writ',
    })
  })

  it('cuts back to the last member boundary when a truncated KEY cannot close', () => {
    expect(salvage('{"language":"de","the')).toEqual({ language: 'de' })
  })

  it('recovers truncated nested objects and array values', () => {
    expect(salvage('{"starPrompt":{"at":1,"b":')).toEqual({ starPrompt: { at: 1 } })
    expect(salvage('{"lastDialogDirs":["a","b","c')).toEqual({ lastDialogDirs: ['a', 'b', 'c'] })
  })

  it('survives doubled separators near the truncation point', () => {
    expect(salvage('{"a":1,,')).toEqual({ a: 1 })
  })

  it('keeps the first complete object when garbage follows it', () => {
    expect(salvage('{"a":1}{"b":2}junk')).toEqual({ a: 1 })
  })

  it('rejects wrong-type, empty and whitespace roots as unsalvageable', () => {
    for (const text of ['', '   \n\t ', '[1, 2, 3]', '42', '"settings"', 'null', 'true']) {
      expect(recoverCorruptAppSettings(text)).toBeNull()
    }
  })

  it('handles a BOM combined with trailing garbage', () => {
    expect(salvage('\uFEFF{"language":"de"} junk')).toEqual({ language: 'de' })
  })

  it('recovers a large pretty-printed file truncated mid-way', () => {
    const entries = Array.from({ length: 5000 }, (_, i) => `  "k${i}": ${i}`).join(',\n')
    const full = `{\n${entries}\n}\n`
    const truncated = full.slice(0, Math.floor(full.length * 0.6))
    const recovered = salvage(truncated)
    expect(recovered).not.toBeNull()
    expect(recovered!.k0).toBe(0)
    expect(Object.keys(recovered!).length).toBeGreaterThan(1000)
  })
})

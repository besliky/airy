import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  AUTHOR_NAME_KEY,
  queueAppSettingsUpdate,
  readAuthorNameSetting,
  sanitizeAuthorName,
} from '@airy-office/electron-utils'
import { readAppSettings, writeAppSetting, writeAppSettings } from '../src/main/app-settings'

/**
 * userData/app-settings.json helpers (src/main/app-settings.ts): a flat JSON
 * object shared by the language preference and the first-run onboarding flag.
 * BUG-1771: the shell wrappers inherit the corrupt-file contract from the
 * single writer — corrupt bytes are preserved as `.bak` before any rewrite
 * and salvaged as far as strict JSON parsing allows.
 */

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'app-settings-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('readAppSettings', () => {
  it('returns an empty object when the file does not exist', () => {
    expect(readAppSettings(settingsPath)).toEqual({})
  })

  it('returns an empty object for invalid JSON', () => {
    writeFileSync(settingsPath, 'not json')
    expect(readAppSettings(settingsPath)).toEqual({})
  })

  it('returns an empty object when the JSON root is not an object', () => {
    writeFileSync(settingsPath, '[1, 2]')
    expect(readAppSettings(settingsPath)).toEqual({})
    writeFileSync(settingsPath, '"zh"')
    expect(readAppSettings(settingsPath)).toEqual({})
  })

  it('parses a valid settings object', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'zh', onboardingSeen: true }))
    expect(readAppSettings(settingsPath)).toEqual({ language: 'zh', onboardingSeen: true })
  })
})

describe('writeAppSetting', () => {
  it('creates the file with the single key on first write', () => {
    writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ onboardingSeen: true })
  })

  it('preserves unrelated existing keys', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    writeAppSetting(settingsPath, 'onboardingSeen', true)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'ja',
      onboardingSeen: true,
    })
  })

  it('overwrites the value of an existing key', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'ja' }))
    writeAppSetting(settingsPath, 'language', 'en')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })

  it('recovers from a corrupt file by rewriting it', () => {
    writeFileSync(settingsPath, '{broken')
    writeAppSetting(settingsPath, 'language', 'en')
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({ language: 'en' })
  })
})

describe('writeAppSettings', () => {
  it('persists onboarding completion and analytics choice together', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'en' }))
    writeAppSettings(settingsPath, { onboardingSeen: true, liveBridge: false })
    expect(JSON.parse(readFileSync(settingsPath, 'utf8'))).toEqual({
      language: 'en',
      onboardingSeen: true,
      liveBridge: false,
    })
  })
})

describe('authorName (settings round-trip)', () => {
  it('persists like the other General settings and reads back sanitized', () => {
    writeAppSetting(settingsPath, AUTHOR_NAME_KEY, sanitizeAuthorName('  Ada \u0007 Lovelace  '))
    expect(readAuthorNameSetting(settingsPath)).toBe('Ada Lovelace')
  })

  it('clearing the name keeps the other keys intact', () => {
    writeAppSettings(settingsPath, { language: 'de', [AUTHOR_NAME_KEY]: 'Ada' })
    writeAppSetting(settingsPath, AUTHOR_NAME_KEY, '')
    const stored = JSON.parse(readFileSync(settingsPath, 'utf8'))
    expect(stored).toEqual({ language: 'de', [AUTHOR_NAME_KEY]: '' })
    expect(readAuthorNameSetting(settingsPath)).toBe('')
  })
})

describe('single writer shared with the editor modules (OBS-1532)', () => {
  it('synchronous shell writes coexist with the queued writer without key loss', async () => {
    // The shell's helpers and the queued writer used by dialog-memory are
    // the SAME module instance (workspace symlink): interleaving both must
    // never drop a key, whatever the interleaving.
    writeAppSetting(settingsPath, 'onboardingSeen', true)
    const queued = Array.from({ length: 40 }, (_, i) =>
      queueAppSettingsUpdate(settingsPath, (current) => ({
        ...current,
        lastDialogDirs: [
          ...(Array.isArray(current.lastDialogDirs) ? current.lastDialogDirs : []),
          i,
        ],
      })),
    )
    for (let i = 0; i < 40; i++) writeAppSetting(settingsPath, `shellKey${i}`, i)
    await Promise.all(queued)
    writeAppSettings(settingsPath, { starPrompt: { resolved: true } })
    const stored = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect(stored.onboardingSeen).toBe(true)
    expect(stored.starPrompt).toEqual({ resolved: true })
    for (let i = 0; i < 40; i++) expect(stored[`shellKey${i}`]).toBe(i)
    expect(stored.lastDialogDirs).toHaveLength(40)
  })
})

// ── BUG-1771: the seven audited corruption variants, through the shell API ──
// SET-26-1: any unreadable byte used to mean total silent loss of ALL
// settings on the next write. Now every variant leaves its exact corrupt
// bytes in `app-settings.json.bak` (forensic recovery for the user), and the
// repairable ones come back through the ordinary merge-write.

const CORRUPT_VARIANTS = [
  {
    name: 'truncated JSON',
    content: '{"language":"de","theme":"dark","autoSaveDef',
    salvaged: { language: 'de', theme: 'dark' },
  },
  {
    name: 'trailing garbage after a valid object',
    content: '{"language":"de","theme":"dark"} trailing garbage',
    salvaged: { language: 'de', theme: 'dark' },
  },
  {
    name: 'UTF-8 BOM before a valid object',
    content: '\uFEFF{"language":"de","theme":"dark"}',
    salvaged: { language: 'de', theme: 'dark' },
  },
  { name: 'empty file', content: '', salvaged: {} },
  { name: 'JSON array root', content: '[1, 2, 3]', salvaged: {} },
  { name: 'JSON number root', content: '42', salvaged: {} },
  { name: 'whitespace-only file', content: '  \n\t ', salvaged: {} },
] as const

describe('BUG-1771 corrupt file handling through the shell wrappers', () => {
  for (const variant of CORRUPT_VARIANTS) {
    it(`${variant.name}: read answers salvaged settings and preserves the bytes`, () => {
      writeFileSync(settingsPath, variant.content)
      expect(readAppSettings(settingsPath)).toEqual(variant.salvaged)
      // forensic copy with the EXACT corrupt bytes, corrupt original untouched
      expect(readFileSync(`${settingsPath}.bak`, 'utf8')).toBe(variant.content)
      expect(readFileSync(settingsPath, 'utf8')).toBe(variant.content)
    })

    it(`${variant.name}: merge-write restores salvaged keys instead of dropping them`, () => {
      writeFileSync(settingsPath, variant.content)
      // the audited repro: one settings toggle must not erase the user's keys
      writeAppSetting(settingsPath, 'starPrompt', { resolved: true })
      const stored = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
      expect(stored.starPrompt).toEqual({ resolved: true })
      for (const [key, value] of Object.entries(variant.salvaged)) {
        expect(stored[key]).toEqual(value)
      }
      expect(readFileSync(`${settingsPath}.bak`, 'utf8')).toBe(variant.content)
    })
  }

  it('a valid file never gets a .bak and reading stays side-effect free', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      writeFileSync(settingsPath, JSON.stringify({ language: 'de', theme: 'dark' }))
      expect(readAppSettings(settingsPath)).toEqual({ language: 'de', theme: 'dark' })
      expect(existsSync(`${settingsPath}.bak`)).toBe(false)
      expect(warn).not.toHaveBeenCalled()
    } finally {
      warn.mockRestore()
    }
  })

  it('a missing file gets no .bak (normal first run)', () => {
    expect(readAppSettings(settingsPath)).toEqual({})
    expect(existsSync(`${settingsPath}.bak`)).toBe(false)
  })

  it('the queued writer (#108) keeps working over a corrupt file', async () => {
    const corrupt = '{"language":"de","theme":"dark","autoSaveDef'
    writeFileSync(settingsPath, corrupt)
    await queueAppSettingsUpdate(settingsPath, (current) => ({
      ...current,
      onboardingSeen: true,
    }))
    expect(readAppSettings(settingsPath)).toEqual({
      language: 'de',
      theme: 'dark',
      onboardingSeen: true,
    })
    expect(readFileSync(`${settingsPath}.bak`, 'utf8')).toBe(corrupt)
  })
})

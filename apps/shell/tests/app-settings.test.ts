import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
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

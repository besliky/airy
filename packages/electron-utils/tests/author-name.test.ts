import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  AUTHOR_NAME_KEY,
  AUTHOR_NAME_MAX,
  configuredAuthorName,
  readAuthorNameSetting,
  sanitizeAuthorName,
} from '../src/author-name'

/**
 * authorName helpers (src/author-name.ts): the display name stamped on new
 * comments / revision marks by every editor main.
 */

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'author-name-'))
  settingsPath = join(dir, 'app-settings.json')
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('sanitizeAuthorName', () => {
  it('passes a plain name through unchanged', () => {
    expect(sanitizeAuthorName('Ada Lovelace')).toBe('Ada Lovelace')
  })

  it('returns an empty string for non-strings', () => {
    expect(sanitizeAuthorName(undefined)).toBe('')
    expect(sanitizeAuthorName(null)).toBe('')
    expect(sanitizeAuthorName(42)).toBe('')
    expect(sanitizeAuthorName({ name: 'Ada' })).toBe('')
  })

  it('drops control characters', () => {
    // controls turn into separators, so words stay readable
    expect(sanitizeAuthorName('Ada\u0000Lovel\u0007ace')).toBe('Ada Lovel ace')
    expect(sanitizeAuthorName('Ada\u007f')).toBe('Ada')
  })

  it('collapses whitespace runs and trims', () => {
    expect(sanitizeAuthorName('  Ada \t\n Lovelace  ')).toBe('Ada Lovelace')
    expect(sanitizeAuthorName('   ')).toBe('')
  })

  it('caps the length at AUTHOR_NAME_MAX', () => {
    const long = 'a'.repeat(AUTHOR_NAME_MAX + 10)
    expect(sanitizeAuthorName(long)).toHaveLength(AUTHOR_NAME_MAX)
    // a name exactly at the cap survives whole
    expect(sanitizeAuthorName('b'.repeat(AUTHOR_NAME_MAX))).toHaveLength(AUTHOR_NAME_MAX)
  })
})

describe('readAuthorNameSetting', () => {
  it('returns an empty string when the file is missing', () => {
    expect(readAuthorNameSetting(settingsPath)).toBe('')
  })

  it('returns an empty string for invalid JSON or a non-object root', () => {
    writeFileSync(settingsPath, 'not json')
    expect(readAuthorNameSetting(settingsPath)).toBe('')
    writeFileSync(settingsPath, '[1]')
    expect(readAuthorNameSetting(settingsPath)).toBe('')
  })

  it('returns an empty string when the key is absent or not a string', () => {
    writeFileSync(settingsPath, JSON.stringify({ language: 'en' }))
    expect(readAuthorNameSetting(settingsPath)).toBe('')
    writeFileSync(settingsPath, JSON.stringify({ [AUTHOR_NAME_KEY]: 7 }))
    expect(readAuthorNameSetting(settingsPath)).toBe('')
  })

  it('reads and sanitizes the stored value, leaving other keys alone', () => {
    writeFileSync(
      settingsPath,
      JSON.stringify({ language: 'ja', [AUTHOR_NAME_KEY]: ' Grace \u0007 Hopper ' }),
    )
    expect(readAuthorNameSetting(settingsPath)).toBe('Grace Hopper')
  })
})

describe('configuredAuthorName', () => {
  it('resolves the setting from the provider userData directory', () => {
    writeFileSync(settingsPath, JSON.stringify({ [AUTHOR_NAME_KEY]: 'Grace Hopper' }))
    const app = { getPath: () => dir }
    expect(configuredAuthorName(app)).toBe('Grace Hopper')
  })

  it('returns an empty string when nothing is configured', () => {
    const app = { getPath: () => dir }
    expect(configuredAuthorName(app)).toBe('')
  })
})

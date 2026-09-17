import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { mainStrings } from '../src/main/i18n/strings-main'

/**
 * Main-process locale shards (src/main/i18n/main/*.ts aggregated by
 * strings-main.ts): zh defines the key set; every other locale must cover
 * exactly the same keys with real content and matching {placeholder} sets.
 * The reverse-usage scan keeps the set free of dead keys.
 */

const locales = Object.keys(mainStrings) as Array<keyof typeof mainStrings>
const referenceKeys = Object.keys(mainStrings.zh).sort()

/** placeholders like {n}, {pages}, {ext} embedded in a template */
function placeholdersOf(template: string): string[] {
  return (template.match(/\{[a-zA-Z0-9]+\}/g) ?? []).sort()
}

describe('main-process locale tables', () => {
  it('includes the expected UI languages', () => {
    expect(locales).toContain('zh')
    expect(locales).toContain('en')
    expect(locales).toContain('zh-TW')
    expect(locales.length).toBeGreaterThanOrEqual(20)
  })

  it.each(locales)('locale %s has exactly the zh key set', (locale) => {
    expect(Object.keys(mainStrings[locale]).sort()).toEqual(referenceKeys)
  })

  it.each(locales)('locale %s has no empty or whitespace-only values', (locale) => {
    const empty = Object.entries(mainStrings[locale]).filter(
      ([, value]) => typeof value !== 'string' || value.trim() === '',
    )
    expect(empty).toEqual([])
  })

  it.each(locales)('locale %s keeps the zh placeholders in every template', (locale) => {
    const table = mainStrings[locale] as Record<string, string>
    const zhTable = mainStrings.zh as Record<string, string>
    const mismatched = referenceKeys.filter(
      (key) => placeholdersOf(table[key]).join(',') !== placeholdersOf(zhTable[key]).join(','),
    )
    expect(mismatched).toEqual([])
  })
})

describe('no dead main keys', () => {
  /** collect non-i18n source files under a directory (usage scan corpus) */
  function collectSource(dir: string, out: string[] = []): string[] {
    for (const entry of readdirSync(dir)) {
      const full = join(dir, entry)
      if (statSync(full).isDirectory()) {
        if (entry === 'i18n') continue
        collectSource(full, out)
      } else if (/\.(ts|tsx)$/.test(entry)) {
        out.push(full)
      }
    }
    return out
  }

  it('every key is referenced outside the i18n shards', () => {
    // grep-equivalent reverse-usage scan (see tests/strings.test.ts); the
    // menu labels the shell builds live in index.ts and the crash prompts
    // in error-dialog.ts — a key nothing reads is dead in all 20 locales.
    const corpus = [
      ...collectSource(join(__dirname, '../src/main')),
      ...collectSource(join(__dirname, '../src/renderer')),
    ]
      .map((file) => readFileSync(file, 'utf8'))
      .join('\n')
    const dead = referenceKeys.filter((key) => !new RegExp(`['"\`]${key}['"\`]`).test(corpus))
    expect(dead).toEqual([])
  })
})

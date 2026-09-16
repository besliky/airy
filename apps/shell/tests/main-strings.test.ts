import { describe, expect, it } from 'vitest'
import { mainStrings } from '../src/main/i18n/strings-main'

/**
 * Main-process locale shards (src/main/i18n/main/*.ts aggregated by
 * strings-main.ts): zh defines the key set; every other locale must cover
 * exactly the same keys with real content and matching {placeholder} sets.
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
    for (const key of referenceKeys) {
      expect(`${key}: ${placeholdersOf(mainStrings[locale][key])}`).toBe(
        `${key}: ${placeholdersOf(mainStrings.zh[key])}`,
      )
    }
  })
})

import { describe, expect, it } from 'vitest'
import { homeStrings } from '../src/renderer/src/i18n/strings-home'
import { settingsStrings } from '../src/renderer/src/i18n/strings-settings'
import { onboardingStrings } from '../src/renderer/src/i18n/strings-onboarding'
import { miscStrings } from '../src/renderer/src/i18n/strings-misc'
import { strings } from '../src/renderer/src/strings'

/**
 * Renderer locale tables. The dictionary is sharded per domain
 * (src/renderer/src/i18n/<domain>/<lang>.ts aggregated by the
 * strings-<domain>.ts files): zh defines each domain's key set; every other
 * locale shard must cover exactly those keys with real content and matching
 * {placeholder} sets. The aggregated table must be the exact union.
 */

const shards = {
  home: homeStrings,
  settings: settingsStrings,
  onboarding: onboardingStrings,
  misc: miscStrings,
}

const locales = Object.keys(strings) as Array<keyof typeof strings>
const referenceKeys = Object.keys(strings.zh).sort()

/** placeholders like {n}, {name}, {v} embedded in a template */
function placeholdersOf(template: string): string[] {
  return (template.match(/\{[a-zA-Z0-9]+\}/g) ?? []).sort()
}

describe('renderer locale shards', () => {
  it('covers every domain with every shard sharing the same locale set', () => {
    const shardLocales = Object.entries(shards).map(
      ([name, dict]) => `${name}:${Object.keys(dict).length}`,
    )
    expect(shardLocales.every((entry) => entry.endsWith(`:${locales.length}`))).toBe(true)
  })

  it.each(Object.entries(shards))(
    'shard %s has exactly the zh key set per locale',
    (_name, dict) => {
      const zhKeys = Object.keys(dict.zh).sort()
      expect(zhKeys.length).toBeGreaterThan(0)
      for (const locale of Object.keys(dict)) {
        expect(Object.keys(dict[locale as keyof typeof dict]).sort(), locale).toEqual(zhKeys)
      }
    },
  )

  it.each(Object.entries(shards))('shard %s has no empty values', (_name, dict) => {
    for (const [locale, table] of Object.entries(dict)) {
      const empty = Object.entries(table).filter(
        ([, value]) => typeof value !== 'string' || value.trim().length === 0,
      )
      expect(empty, `${locale} has empty values`).toEqual([])
    }
  })

  it('aggregates to exactly the union of the shard key sets', () => {
    const union = Object.values(shards).flatMap((dict) => Object.keys(dict.zh))
    expect(union.length).toBe(new Set(union).size) // no cross-shard duplicates
    expect([...union].sort()).toEqual(referenceKeys)
  })
})

describe('home-screen locale tables', () => {
  it('includes the expected UI languages', () => {
    expect(locales).toContain('zh')
    expect(locales).toContain('en')
    expect(locales).toContain('zh-TW')
    expect(locales.length).toBeGreaterThanOrEqual(19)
  })

  it.each(locales)('locale %s has exactly the zh key set', (locale) => {
    expect(Object.keys(strings[locale]).sort()).toEqual(referenceKeys)
  })

  it.each(locales)('locale %s has no empty or whitespace-only values', (locale) => {
    const empty = Object.entries(strings[locale])
      .filter(([, value]) => typeof value !== 'string' || value.trim().length === 0)
      .map(([key]) => key)
    expect(empty).toEqual([])
  })

  it.each(locales)('locale %s keeps the zh placeholder set for each key', (locale) => {
    const table = strings[locale] as Record<string, string>
    const mismatched = referenceKeys.filter((key) => {
      const localePlaceholders = placeholdersOf(table[key])
      const zhPlaceholders = placeholdersOf((strings.zh as Record<string, string>)[key])
      // Singular-count keys ("...One") may drop the numeral entirely in
      // languages that express "one" grammatically (e.g. ar/he), so a
      // missing placeholder is fine there — an extra one is not.
      if (key.endsWith('One')) {
        return localePlaceholders.some((p) => !zhPlaceholders.includes(p))
      }
      return localePlaceholders.join(',') !== zhPlaceholders.join(',')
    })
    expect(mismatched).toEqual([])
  })

  it('has no duplicate values that suggest an untranslated copy-paste between zh and en', () => {
    // sanity check that en is actually translated, not a zh copy
    const zh = strings.zh as Record<string, string>
    const en = strings.en as Record<string, string>
    const identical = referenceKeys.filter((key) => zh[key] === en[key])
    // a few shared strings (brand names, "PDF", "OK"-style tokens) are fine
    expect(identical.length).toBeLessThan(referenceKeys.length / 4)
  })
})

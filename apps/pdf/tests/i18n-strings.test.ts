import { describe, expect, it } from 'vitest'
import { LANGS } from '@airy-office/i18n'
import { strings } from '../src/renderer/i18n/strings'
import {
  fillFormStrings,
  fillFormStringsFor,
  localizedFillFormStrings,
} from '../src/renderer/i18n/strings-fillform'

const dicts = strings as Record<string, Record<string, string>>
const zhKeys = Object.keys(dicts.zh!).sort()

describe('i18n string tables', () => {
  it('provides a dictionary for every supported language and nothing else', () => {
    expect(Object.keys(dicts).sort()).toEqual([...LANGS].sort())
  })

  it.each([...LANGS])('locale %s has exactly the zh key set', (lang) => {
    expect(Object.keys(dicts[lang]!).sort()).toEqual(zhKeys)
  })

  it.each([...LANGS])('locale %s has no empty values', (lang) => {
    for (const [key, value] of Object.entries(dicts[lang]!)) {
      expect(typeof value, `${lang}.${key}`).toBe('string')
      expect(value.trim().length, `${lang}.${key} is empty`).toBeGreaterThan(0)
    }
  })

  it.each([...LANGS])('locale %s keeps the same placeholders as zh', (lang) => {
    for (const key of zhKeys) {
      const placeholders = (s: string) => (s.match(/\{\w+\}/g) ?? []).sort()
      expect(placeholders(dicts[lang]![key]!), `${lang}.${key}`).toEqual(
        placeholders(dicts.zh![key]!),
      )
    }
  })
})

describe('fill-form locale shards', () => {
  // this domain differs from the zh-defines-the-key-set pattern: en is the
  // fallback base, and a locale shard may omit keys (zh and zh-TW omit the
  // locale-neutral formFieldProgress) — the lookup then fills them from en
  const enKeys = Object.keys(fillFormStrings).sort()
  const shardLocales = Object.keys(localizedFillFormStrings)

  it('has an en base shard plus one shard per every other supported language', () => {
    expect(['en', ...shardLocales].sort()).toEqual([...LANGS].sort())
  })

  it.each(shardLocales)('shard %s holds only keys the en base defines', (locale) => {
    const keys = Object.keys(
      localizedFillFormStrings[locale as keyof typeof localizedFillFormStrings],
    )
    expect(
      keys.filter((key) => !enKeys.includes(key)),
      locale,
    ).toEqual([])
  })

  it.each(shardLocales)('lookup %s resolves exactly the en key set', (locale) => {
    expect(Object.keys(fillFormStringsFor(locale)).sort()).toEqual(enKeys)
  })
})

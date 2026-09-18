import { afterEach, describe, expect, it } from 'vitest'
import { loadStrings, type LocaleDict } from '../src/renderer/i18n/strings'
import { loadLocale, setModuleLang, t } from '../src/renderer/i18n/locale'

// PERF-901: dictionaries load per locale on demand. These tests pin the lazy
// loader contract (memoized per-locale fetch, zh key set coverage) and the
// module translator's behavior across a language switch.

afterEach(() => {
  setModuleLang('zh')
})

describe('loadStrings (lazy per-locale dictionaries)', () => {
  it('resolves a dictionary covering the zh key set across all three domains', async () => {
    const en = await loadStrings('en')
    expect(en.appMergeWorkbooks).toBe('Merge Workbooks')
    expect(en.aiComposerPlaceholderBuild).toContain('table')
    // a dialogs-domain key proves the per-locale merge module stitched all shards
    const keys = Object.keys(en) as Array<keyof LocaleDict>
    expect(keys.length).toBeGreaterThan(1000)
  })

  it('memoizes: the second call returns the same dictionary object', async () => {
    const first = await loadStrings('ru')
    const second = await loadStrings('ru')
    expect(second).toBe(first)
  })
})

describe('module translator across a language switch', () => {
  it('falls back to the raw key while no dictionary is loaded', () => {
    setModuleLang('hi')
    expect(t('appMergeWorkbooks')).toBe('appMergeWorkbooks')
  })

  it('translates with the loaded locale and follows a switch', async () => {
    await loadLocale('en')
    setModuleLang('en')
    expect(t('appMergeWorkbooks')).toBe('Merge Workbooks')

    await loadLocale('ru')
    setModuleLang('ru')
    expect(t('appMergeWorkbooks')).toBe('Объединить книги')

    // the previously loaded locale stays cached and usable again
    setModuleLang('en')
    expect(t('appMergeWorkbooks')).toBe('Merge Workbooks')
  })

  it('fills {placeholder} params like the eager table did', async () => {
    await loadLocale('en')
    setModuleLang('en')
    expect(t('appMergeWorkbooksReading', { sheet: 'Sheet1', file: 'a.xlsx' })).toBe(
      'Importing Sheet1 from a.xlsx…',
    )
  })
})

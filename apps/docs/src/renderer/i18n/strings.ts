import type { Lang } from '@airy-office/i18n'
import type { aiStrings } from './strings-ai'
import type { appStrings } from './strings-app'
import type { editorStrings } from './strings-editor'
import type { layoutStrings } from './strings-layout'
import type { referencesStrings } from './strings-references'
import type { reviewStrings } from './strings-review'
import type { ribbonStrings } from './strings-ribbon'
import type { tableStrings } from './strings-table'

/// PERF-904: the 20-locale dictionary used to be one eager blob inside the
/// entry chunk (locale.tsx statically imported the merged `strings` object).
/// It is data, not logic, so it now loads per locale on demand:
/// `locales/<lang>.ts` merges the eight domain shards for one language and
/// each becomes its own lazy chunk, fetched by `loadStrings` before the first
/// render (in main.tsx's bootstrap) and before a language switch commits.
/// The aggregators above stay in the type graph only — `import type` is
/// erased at runtime, so no dictionary ships with the entry chunk and the
/// sharding contract (zh defines the key set) keeps compile-time enforcement.

/** merged dictionary for a single locale; the zh shards define the key set */
export type LocaleDict = (typeof appStrings)['zh'] &
  (typeof ribbonStrings)['zh'] &
  (typeof referencesStrings)['zh'] &
  (typeof tableStrings)['zh'] &
  (typeof editorStrings)['zh'] &
  (typeof reviewStrings)['zh'] &
  (typeof aiStrings)['zh'] &
  (typeof layoutStrings)['zh']

/** one dynamic-import entry per locale — Rollup turns each into a lazy chunk */
const LOADERS: Record<Lang, () => Promise<{ default: LocaleDict }>> = {
  zh: () => import('./locales/zh'),
  en: () => import('./locales/en'),
  ja: () => import('./locales/ja'),
  ko: () => import('./locales/ko'),
  fr: () => import('./locales/fr'),
  de: () => import('./locales/de'),
  es: () => import('./locales/es'),
  th: () => import('./locales/th'),
  id: () => import('./locales/id'),
  ru: () => import('./locales/ru'),
  ar: () => import('./locales/ar'),
  pt: () => import('./locales/pt'),
  it: () => import('./locales/it'),
  pl: () => import('./locales/pl'),
  cs: () => import('./locales/cs'),
  nl: () => import('./locales/nl'),
  ms: () => import('./locales/ms'),
  he: () => import('./locales/he'),
  hi: () => import('./locales/hi'),
  'zh-TW': () => import('./locales/zh-TW'),
}

const cache = new Map<Lang, LocaleDict>()

/** fetch (and memoize) the dictionary of one locale */
export function loadStrings(lang: Lang): Promise<LocaleDict> {
  const cached = cache.get(lang)
  if (cached) return Promise.resolve(cached)
  return LOADERS[lang]().then((module) => {
    const dict = module.default
    cache.set(lang, dict)
    return dict
  })
}

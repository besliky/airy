import type { Lang } from '@airy-office/i18n'
import type { aiStrings } from './strings-ai'
import type { appStrings } from './strings-app'
import type { dialogStrings } from './strings-dialogs'

/// PERF-901: the 19-locale dictionary used to be one eager `app-i18n` chunk
/// (~2.1 MB) because `locale.tsx` statically imported the merged object. It is
/// data, not logic, so it now loads per locale on demand: `locales/<lang>.ts`
/// merges the three domain shards for one language and each becomes its own
/// lazy chunk, fetched by `loadStrings` before the first render (in parallel
/// with the cell-font preload in main.tsx) and before a language switch
/// commits. The aggregators below stay in the type graph only — `import type`
/// is erased at runtime, so no dictionary ships with the entry chunk and the
/// sharding contract (zh defines the key set) keeps compile-time enforcement.

/** merged dictionary for a single locale; the zh shards define the key set */
export type LocaleDict = (typeof appStrings)['zh'] &
  (typeof dialogStrings)['zh'] &
  (typeof aiStrings)['zh']

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

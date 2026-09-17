import { ar } from './fillform/ar'
import { cs } from './fillform/cs'
import { de } from './fillform/de'
import { en } from './fillform/en'
import { es } from './fillform/es'
import { fr } from './fillform/fr'
import { he } from './fillform/he'
import { hi } from './fillform/hi'
import { id } from './fillform/id'
import { it } from './fillform/it'
import { ja } from './fillform/ja'
import { ko } from './fillform/ko'
import { ms } from './fillform/ms'
import { nl } from './fillform/nl'
import { pl } from './fillform/pl'
import { pt } from './fillform/pt'
import { ru } from './fillform/ru'
import { th } from './fillform/th'
import { zh } from './fillform/zh'
import { zhTW } from './fillform/zh-TW'

/** Fill Form ribbon strings, sharded per locale (i18n/fillform/<lang>.ts).
 * Unlike the other domains, en is the fallback base here and defines the
 * full key set: a locale shard may omit keys (zh and zh-TW omit
 * formFieldProgress, whose "{current} / {total}" is locale-neutral) and the
 * gap is filled from en at lookup time. The Partial<Record<keyof typeof
 * en, string>> check on each shard still rejects unknown keys. */
/** en base (the fallback for keys a locale shard omits). */
export const fillFormStrings = en

export const localizedFillFormStrings = {
  zh,
  'zh-TW': zhTW,
  cs,
  ja,
  ko,
  fr,
  de,
  es,
  th,
  id,
  ru,
  ar,
  pt,
  it,
  pl,
  nl,
  ms,
  he,
  hi,
} as const

export const fillFormStringsFor = (lang: string) => ({
  ...fillFormStrings,
  ...(localizedFillFormStrings[lang as keyof typeof localizedFillFormStrings] ?? {}),
})

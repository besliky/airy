import { defineStrings } from '@airy-office/i18n'
import { zh } from './misc/zh'
import { en } from './misc/en'
import { ja } from './misc/ja'
import { ko } from './misc/ko'
import { fr } from './misc/fr'
import { de } from './misc/de'
import { es } from './misc/es'
import { th } from './misc/th'
import { id } from './misc/id'
import { ru } from './misc/ru'
import { ar } from './misc/ar'
import { pt } from './misc/pt'
import { it } from './misc/it'
import { pl } from './misc/pl'
import { cs } from './misc/cs'
import { nl } from './misc/nl'
import { ms } from './misc/ms'
import { he } from './misc/he'
import { hi } from './misc/hi'
import { zhTW } from './misc/zh-TW'

/** Tab bar, error and relative-time strings; zh defines the key set, siblings must match (type-checked). */
export const miscStrings = defineStrings({
  zh,
  en,
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
  cs,
  nl,
  ms,
  he,
  hi,
  'zh-TW': zhTW,
})

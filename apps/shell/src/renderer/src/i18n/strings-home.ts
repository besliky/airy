import { defineStrings } from '@airy-office/i18n'
import { zh } from './home/zh'
import { en } from './home/en'
import { ja } from './home/ja'
import { ko } from './home/ko'
import { fr } from './home/fr'
import { de } from './home/de'
import { es } from './home/es'
import { th } from './home/th'
import { id } from './home/id'
import { ru } from './home/ru'
import { ar } from './home/ar'
import { pt } from './home/pt'
import { it } from './home/it'
import { pl } from './home/pl'
import { cs } from './home/cs'
import { nl } from './home/nl'
import { ms } from './home/ms'
import { he } from './home/he'
import { hi } from './home/hi'
import { zhTW } from './home/zh-TW'

/** Home screen strings (file lists, projects, quick start, account); zh defines the key set, siblings must match (type-checked). */
export const homeStrings = defineStrings({
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

import { defineStrings } from '@airy-office/i18n'
import { zh } from './main/zh'
import { en } from './main/en'
import { ja } from './main/ja'
import { ko } from './main/ko'
import { fr } from './main/fr'
import { de } from './main/de'
import { es } from './main/es'
import { th } from './main/th'
import { id } from './main/id'
import { ru } from './main/ru'
import { ar } from './main/ar'
import { pt } from './main/pt'
import { it } from './main/it'
import { pl } from './main/pl'
import { cs } from './main/cs'
import { nl } from './main/nl'
import { ms } from './main/ms'
import { he } from './main/he'
import { hi } from './main/hi'
import { zhTW } from './main/zh-TW'

/** slides main-process strings (dialogs, native menus, export, autosave prompts); zh defines the key set every other locale shard must match */
export const mainStrings = defineStrings({
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

import { defineStrings } from '@airy-office/i18n'
import { zh } from './onboarding/zh'
import { en } from './onboarding/en'
import { ja } from './onboarding/ja'
import { ko } from './onboarding/ko'
import { fr } from './onboarding/fr'
import { de } from './onboarding/de'
import { es } from './onboarding/es'
import { th } from './onboarding/th'
import { id } from './onboarding/id'
import { ru } from './onboarding/ru'
import { ar } from './onboarding/ar'
import { pt } from './onboarding/pt'
import { it } from './onboarding/it'
import { pl } from './onboarding/pl'
import { cs } from './onboarding/cs'
import { nl } from './onboarding/nl'
import { ms } from './onboarding/ms'
import { he } from './onboarding/he'
import { hi } from './onboarding/hi'
import { zhTW } from './onboarding/zh-TW'

/** First-run onboarding strings; zh defines the key set, siblings must match (type-checked). */
export const onboardingStrings = defineStrings({
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

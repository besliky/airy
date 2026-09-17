/** Main-process i18n strings for Airy Slides (dialogs, native menus, export, autosave prompts). */
import { createI18n, getUiLang } from '@airy-office/i18n'
import { mainStrings } from './i18n/strings-main'

export const tMain = createI18n(mainStrings)

/** Translate with the current UI language. */
export const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(getUiLang(), key, params)

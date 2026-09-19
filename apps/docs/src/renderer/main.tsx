import { createRoot } from 'react-dom/client'
import { htmlDir, htmlLang, type Lang } from '@airy-office/i18n'
import { App } from './App'
import { LocaleProvider, loadLocale, setModuleLang } from './i18n/locale'
import type { UiTheme } from '../shared/ipc'
import '@airy-office/ui/tokens.css'
import '@airy-office/ui/screentip.css'
import '@airy-office/ui/color-picker.css'
import '@airy-office/ui/dropdown.css'
import '@airy-office/ui/ribbon-collapse.css'
import '@airy-office/ui/markdown.css'
import '@airy-office/ui/ai-panel-prefs.css'
import '@airy-office/ui/ai-scope-quote.css'
import './styles.css'
import './fonts/fonts.css'
import { applyAiPanelPrefs, installScreenTips } from '@airy-office/ui'

installScreenTips()

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

async function bootstrap(): Promise<void> {
  let lang: Lang = 'zh'
  let theme: UiTheme = 'system'
  try {
    // per-promise catch: standalone runs have no app:get-theme handler, and
    // that rejection must not drop a resolved language
    ;[lang, theme] = await Promise.all([
      window.desktop.getLanguage().catch(() => 'zh' as const),
      window.desktop.getTheme().catch(() => 'system' as const),
    ])
  } catch {
    /* dev renderer without the preload bridge */
  }
  // PERF-904: the UI-language dictionary is a lazy chunk fetched for this one
  // locale (the other 19 stay unloaded); it rides next to the theme/IPC setup
  // so it adds no serial wait before the first frame.
  const stringsReady = loadLocale(lang)
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)
  applyTheme(theme)
  await stringsReady
  window.desktop?.onThemeChanged(applyTheme)
  void window.desktop
    ?.getAiPanelPrefs?.()
    .then(applyAiPanelPrefs)
    .catch(() => {})
  window.desktop?.onAiPanelPrefsChanged?.(applyAiPanelPrefs)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()

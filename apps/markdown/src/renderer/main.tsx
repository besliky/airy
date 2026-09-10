import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@airy-office/i18n'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import type { UiTheme } from '../shared/ipc'
import '@airy-office/ui/tokens.css'
import '@airy-office/ui/screentip.css'
import '@airy-office/ui/dropdown.css'
import '@airy-office/ui/find-panel.css'
import '@airy-office/ui/ribbon-collapse.css'
import '@airy-office/ui/markdown.css'
import '@airy-office/ui/ai-panel-prefs.css'
import '@airy-office/ui/ai-scope-quote.css'
import 'katex/dist/katex.min.css'
import './styles.css'
import { applyAiPanelPrefs, installScreenTips } from '@airy-office/ui'

installScreenTips()

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

void (async () => {
  const [lang, theme] = await Promise.all([
    window.markdownApi.getLanguage().catch(() => 'zh' as const),
    window.markdownApi.getTheme().catch(() => 'system' as const),
  ])
  document.documentElement.lang = htmlLang(lang as Lang)
  applyTheme(theme)
  window.markdownApi.onThemeChanged(applyTheme)
  void window.markdownApi
    ?.getAiPanelPrefs?.()
    .then(applyAiPanelPrefs)
    .catch(() => {})
  window.markdownApi?.onAiPanelPrefsChanged?.(applyAiPanelPrefs)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
})()

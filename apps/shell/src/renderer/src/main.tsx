import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlDir, htmlLang } from '@airy-office/i18n'
import airyMark from '@airy-office/ui/assets/airy-mark.png'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import '@airy-office/ui/tokens.css'
import '@airy-office/ui/screentip.css'
import '@airy-office/ui/dropdown.css'
import './home.css'
import './tabbar.css'
import { installScreenTips } from '@airy-office/ui'

installScreenTips()

// macOS shell window is created with vibrancy; a transparent body lets the
// editor views' translucent regions (e.g. slides thumbnail pane) show it
if (navigator.platform.toLowerCase().includes('mac')) document.body.classList.add('vib')

/**
 * First-paint skeleton: the shell renderer renders nothing until the
 * language/onboarding/theme IPC round-trips resolve, leaving a blank beat on
 * slow disks. This paints the shell surface immediately; the spinner only
 * becomes visible after a short delay, so a fast resolve never flashes it.
 */
function StartupSkeleton() {
  return (
    <div className="startup-skeleton" aria-hidden="true">
      <img className="startup-logo" src={airyMark} alt="" />
      <span className="startup-spinner" />
    </div>
  )
}

const root = createRoot(document.getElementById('root')!)
root.render(<StartupSkeleton />)

// resolve the persisted language, first-run flag, and theme before first paint
// so the UI never flashes (home showing briefly before the onboarding overlay)
void Promise.all([
  window.aiOffice.getLanguage(),
  // if the flag is unreadable, skip onboarding rather than block the home screen
  window.aiOffice.onboardingSeen().catch(() => true),
  window.aiOffice.getTheme().catch(() => 'system' as const),
]).then(([lang, onboardingSeen, theme]) => {
  document.documentElement.lang = htmlLang(lang)
  document.documentElement.dir = htmlDir(lang)
  // apply theme attribute before first paint to avoid flash
  if (theme !== 'system') {
    document.documentElement.setAttribute('data-theme', theme)
  }
  window.aiOffice.onThemeChanged((next) => {
    if (next === 'system') document.documentElement.removeAttribute('data-theme')
    else document.documentElement.setAttribute('data-theme', next)
  })
  root.render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        <AppFrame initialOnboardingSeen={onboardingSeen} />
      </LocaleProvider>
    </React.StrictMode>,
  )
})

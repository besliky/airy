/**
 * CSS hyphenation support probe (BUG-1541).
 *
 * doc-style-css emits `hyphens:auto` for settings.xml w:autoHyphenation, but
 * Chromium only hyphenates when it carries hyphenation data for the element's
 * language. Stock Electron builds ship without those dictionaries (audit
 * 2026-09-21: Electron 43.3 / Chromium 150.0.7871.212 hyphenates nothing — a
 * canary with `hyphens:auto`, lang=en-US and a long word produced zero hyphen
 * characters), and there is no switch or API to supply them: the dictionaries
 * are compiled into the Chromium binary or delivered by Chrome's component
 * updater, which Electron does not wire (and Hunspell spellcheck dictionaries
 * are unrelated — `hyphens` never reads them).
 *
 * Rather than emit a declaration that silently does nothing — and could one
 * day start hyphenating by non-Word rules (no hyphenationZone, engine-blind)
 * when an Electron upgrade bundles dictionaries — the renderer probes once:
 * a narrow box holding one long English word must wrap into several line
 * boxes when hyphenation works and stay on a single overflowing line when it
 * does not. Callers pass the result to docStyleCss, which downgrades auto to
 * manual (keeping w:softHyphen / U+00AD break opportunities, exactly what
 * auto achieves without dictionaries).
 */

let supported: boolean | undefined

/** whether CSS hyphens:auto actually hyphenates in this build; cached after the first call */
export function cssHyphenationSupported(): boolean {
  if (supported === undefined) supported = probeCssHyphenation()
  return supported
}

function probeCssHyphenation(): boolean {
  if (typeof document === 'undefined') return false
  const probe = document.createElement('div')
  try {
    probe.lang = 'en-US'
    probe.textContent = 'internationalization'
    // fixed and far off-screen so the probe never paints; word-break and
    // overflow-wrap pinned so an unhyphenated word can only overflow on one
    // line, never break on its own
    probe.style.cssText =
      'position:fixed;left:-10000px;top:0;width:40px;font:12px sans-serif;' +
      'hyphens:auto;-webkit-hyphens:auto;word-break:normal;overflow-wrap:normal'
    document.body.appendChild(probe)
    const range = document.createRange()
    range.selectNodeContents(probe)
    const lines = range.getClientRects().length
    // one overflowing line = no dictionaries; several stacked line boxes = hyphens
    return lines > 1
  } catch {
    // no layout engine (tests, SSR-ish contexts) — treat as unsupported
    return false
  } finally {
    probe.remove()
  }
}

/** test hook: forget the cached probe result so the next call re-probes */
export function resetCssHyphenationProbe(): void {
  supported = undefined
}

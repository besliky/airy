import { app } from 'electron'

/**
 * Lazy accessibility policy (PERF-1700c).
 *
 * Chromium enables its accessibility support on its own when an assistive
 * technology (AT) client is detected: AT-SPI on Linux, UIA/MSAA on Windows,
 * NSAccessibility on macOS. Empirically verified on Linux with Electron 43:
 * an Orca-style AT-SPI client that queries objects (get_attributes /
 * get_relation_set — Orca's signature calls) flips support on at runtime,
 * exposing `nativeAPIs`, `webContents` and `extendedProperties`. Merely
 * connecting to the accessibility bus or walking the tree does NOT flip it,
 * so users without assistive tech never pay for it.
 *
 * Forcing support on for every user (`setAccessibilitySupportEnabled(true)`,
 * the previous behavior) made every renderer build and maintain the AX tree
 * even with no AT present: the tree update runs in the BeginMainFrame commit
 * (`LocalFrameView::RunAccessibilitySteps`) and scaled to ~2s per keystroke
 * on a multi-megabyte paragraph (PERF-1700b). It also enables
 * `inlineTextBoxes` (per-character boxes) which lazy detection does not.
 *
 * Note: on Linux `app.isAccessibilitySupportEnabled()` stays `false` even
 * after lazy detection has enabled support (Electron 43), and the
 * `accessibility-support-changed` event does not fire for that path — so
 * neither can be used to gate anything here; we simply stay out of the way.
 *
 * Escape hatch: AIRY_FORCE_A11Y=1 restores the old always-on behavior for
 * setups whose assistive tech is not detected by Chromium.
 */
export function applyAccessibilityPolicy(): 'forced' | 'detection' {
  if (process.env.AIRY_FORCE_A11Y === '1') {
    app.setAccessibilitySupportEnabled(true)
    return 'forced'
  }
  // Leave Chromium's default: support turns on when an AT client is detected.
  return 'detection'
}

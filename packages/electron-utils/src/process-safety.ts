/**
 * Process-level safety helpers shared by every editor main process. All six
 * apps' main code runs inside the shell's single process, so one rejected
 * promise or crashed renderer must never escalate into losing every open
 * document tab.
 */

/**
 * Fire-and-forget a renderer load (loadURL/loadFile) without risking an
 * unhandled rejection: the promise is logged when it rejects instead of
 * bubbling into process-level crash handling.
 */
export function voidLoad(load: Promise<unknown>, label: string): void {
  load.catch((err: unknown) => {
    console.error(`[load] ${label} failed:`, err)
  })
}

/**
 * Whether a render-process-gone reason should trigger crash recovery.
 * 'oom' and 'crashed' leave a blank zombie tab the user must recover from;
 * 'clean-exit'/'killed' are intentional teardown (view removed, app quitting)
 * and must not prompt.
 */
export function isRecoverableRendererCrash(reason: string): boolean {
  return reason === 'oom' || reason === 'crashed'
}

/**
 * Minimal in-tab error page shown while a crashed renderer awaits the user's
 * Reload/Close decision. Pure data: URL so it works for any webContents
 * without touching the network or the app bundle.
 */
export function crashErrorPageUrl(message: string): string {
  const html = `<!doctype html><html><head><meta charset="utf-8"><style>
html,body{height:100%;margin:0}
body{display:flex;align-items:center;justify-content:center;background:#f5f5f5;color:#333;
font:14px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif;text-align:center;padding:24px}
</style></head><body><div>${message.replace(/[<>&]/g, (c) =>
    c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&amp;',
  )}</div></body></html>`
  return `data:text/html;charset=utf-8,${encodeURIComponent(html)}`
}

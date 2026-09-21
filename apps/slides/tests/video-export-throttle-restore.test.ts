/**
 * BUG-1220 (audit BUG-1210, fixed with PR #83): a video export suspends the
 * exporting webContents' background throttling (MediaRecorder timestamps
 * frames by the wall clock), and the renderer's `finally` was the ONLY
 * restore — a renderer that died mid-recording (OOM on a big deck, encoder
 * crash, tab teardown) never ran it, and the webContents survives a crash
 * (crash recovery reloads into the same one), leaving its timers unclamped
 * for the rest of the tab's life.
 *
 * The main side now owns the restore: every active export is tracked in
 * videoExportSuspendByWc and reverts on render-process-gone / destroyed,
 * and a fresh set-video-export-active detaches the previous watch first
 * (normal completion path). The freeze-recovery reload also lands here —
 * it goes through forcefullyCrashRenderer(), which fires
 * render-process-gone.
 *
 * Source-wiring test (same style as export-render-cancel / the shell's
 * generated-open-sender-routing): slides-main.ts is an Electron main module
 * that cannot be imported into a unit test, so the lifecycle wiring is the
 * contract under regression guard.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const source = readFileSync(join(here, '../src/main/slides-main.ts'), 'utf8')

describe('video-export throttling suspend cannot outlive the export (BUG-1220)', () => {
  it('main tracks active exports per webContents', () => {
    expect(source).toContain('const videoExportSuspendByWc = new Map<number, () => void>()')
  })

  it('an activated export arms restore on both renderer-death paths', () => {
    // renderer crash (which the freeze-recovery reload also goes through via
    // forcefullyCrashRenderer) and webContents teardown each restore the
    // throttling from main, because the renderer's finally never runs there
    expect(source).toContain("wc.once('render-process-gone', restore)")
    expect(source).toContain("wc.once('destroyed', restore)")
  })

  it('restore re-enables throttling guarded against a dead webContents', () => {
    expect(source).toContain('if (!wc.isDestroyed()) wc.setBackgroundThrottling(true)')
  })

  it('a new set-video-export-active detaches the previous watch (normal completion)', () => {
    expect(source).toContain('videoExportSuspendByWc.get(wc.id)?.()')
    expect(source).toContain('videoExportSuspendByWc.delete(wc.id)')
  })

  it('the only throttling toggle in the app main is the handler itself', () => {
    // the handler's own two calls (set + guarded restore in this file); no
    // other main-side code may touch the flag
    expect(source.match(/setBackgroundThrottling/g)).toHaveLength(2)
  })
})

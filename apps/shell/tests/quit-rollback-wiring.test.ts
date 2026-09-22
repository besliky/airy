/**
 * Source-wiring test for the quit/window-close rollback paths (same style as
 * generated-open-sender-routing.test.ts): the coordinator modules are
 * unit-tested in isolation, but their glue inside the Electron main
 * entrypoints cannot be imported into a unit test — assert the load-bearing
 * wiring lines instead.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const shellMain = readFileSync(join(here, '../src/main/index.ts'), 'utf8')
const sheetsMain = readFileSync(join(here, '../../sheets/src/main/sheets-main.ts'), 'utf8')
const updaterMain = readFileSync(join(here, '../src/main/updater/index.ts'), 'utf8')

describe('quit rollback wiring (BUG-1216)', () => {
  it('before-quit arms quit flow and the rollback-able shutdown effects', () => {
    expect(shellMain).toContain('quitFlow.begin()')
    expect(shellMain).toContain('shutdownEffects.arm()')
    // the effects are no longer armed inline (an inline arming would bypass
    // the coordinator and stay armed after a cancelled quit)
    expect(shellMain).not.toContain('  markSheetsShuttingDown()\n')
  })

  it('abortAppQuit unwinds the shutdown effects before the recovery persist', () => {
    // ordinary-close cancels must not run the quit rollback at all
    expect(shellMain).toContain('if (!quitFlow.quitting) return')
    expect(shellMain).toContain('shutdownEffects.rollback()')
    // ordering: rollback precedes the recovery session write
    const rollbackAt = shellMain.indexOf('shutdownEffects.rollback()')
    const recoveryAt = shellMain.indexOf('if (restoreSession) persistSessionState(false)')
    expect(rollbackAt).toBeGreaterThan(-1)
    expect(recoveryAt).toBeGreaterThan(rollbackAt)
  })

  it('the sheets shutdown flag can be cleared and the listener stays arrow-wrapped', () => {
    // the flag function takes the direction; a bare app.on('before-quit',
    // markSheetsShuttingDown) would store the event object as the flag
    expect(sheetsMain).toContain('export function markSheetsShuttingDown(shuttingDown = true)')
    expect(sheetsMain).toContain("app.on('before-quit', () => markSheetsShuttingDown())")
  })

  it("every shell window's close event runs through the extracted guard (BUG-1217)", () => {
    // the window wiring delegates to the latch-guarded state machine instead
    // of an inline handler that a repeat close event could re-enter
    expect(shellMain).toContain("win.on(\n    'close',\n    createWindowCloseGuard({")
    expect(shellMain).toContain('tagGuardTabs(')
    expect(shellMain).not.toContain('let closeConfirmed = false')
  })

  it('a confirmed-closing window is marked only after its own snapshot write (BUG-1218)', () => {
    // session writes serialize the registry's persistable entries...
    expect(shellMain).toContain('shellWindows.persistableEntries(exclude)')
    // ...and finishWindowClose marks the closer AFTER the persist block: the
    // quit snapshot and the last-window close legitimately include the
    // closer, every LATER write (recovery rewrite, a sibling's ordinary
    // close) must not resurrect it
    const persistAt = shellMain.indexOf(': persistSessionState(decision.skipStaged)')
    const markAt = shellMain.indexOf('entry.closingConfirmed = true')
    expect(persistAt).toBeGreaterThan(-1)
    expect(markAt).toBeGreaterThan(persistAt)
  })
})

describe('quit snapshot write failure re-arm (BUG-1224)', () => {
  it('finishWindowClose resets persist-once when the snapshot write fails', () => {
    // persistSessionState reports whether the write landed (it still logs
    // and never throws); a failed quit-time write must disarm persist-once
    // so the NEXT confirmed close retries instead of leaving a stale
    // session that resurrects already-closed windows
    expect(shellMain).toContain('const written = decision.excludeClosing')
    expect(shellMain).toContain('if (!written) quitFlow.markSnapshotWriteFailed()')
  })
})

describe('before-quit session flush (BUG-408)', () => {
  it('before-quit lands a full session snapshot before the windows start closing', () => {
    // the debounced tab-change save is 800 ms deep: a crash once the quit has
    // started used to lose the last tab changes; the flush must run while
    // every window is still registered (a full snapshot can only widen the
    // persisted set) and before the shutdown effects arm
    const start = shellMain.indexOf("app.on('before-quit'")
    const beforeQuit = shellMain.slice(start)
    expect(beforeQuit).toContain('persistSessionState(false)')
    const persistAt = beforeQuit.indexOf('persistSessionState(false)')
    const beginAt = beforeQuit.indexOf('quitFlow.begin()')
    const armAt = beforeQuit.indexOf('shutdownEffects.arm()')
    expect(beginAt).toBeGreaterThan(-1)
    expect(persistAt).toBeGreaterThan(beginAt)
    expect(armAt).toBeGreaterThan(persistAt)
  })
})

describe('updater install behind the close flow (BUG-411)', () => {
  it('the updater glue requests the install via app.quit, not quitAndInstall directly', () => {
    expect(updaterMain).toContain('pendingInstaller.request()')
    const requestAt = updaterMain.indexOf('pendingInstaller.request()')
    const quitAt = updaterMain.indexOf('app.quit()', requestAt)
    expect(quitAt).toBeGreaterThan(-1)
    // the direct call must be gone: it spawned the NSIS/AppImage installer
    // before any window's dirty guard could object
    expect(updaterMain.match(/quitAndInstall: \(\) => autoUpdater\.quitAndInstall/)).toBeNull()
  })

  it('the install fires only from the flush points (windows gone, no guard left)', () => {
    expect(updaterMain).toContain(
      'if (pendingInstaller.flush()) autoUpdater.quitAndInstall(true, true)',
    )
    // flush point 1: window-all-closed, before the platform quit
    const allClosed = shellMain.indexOf("app.on('window-all-closed'")
    expect(allClosed).toBeGreaterThan(-1)
    expect(shellMain.slice(allClosed)).toContain('flushPendingInstall()')
    const flushAt = shellMain.indexOf('flushPendingInstall()', allClosed)
    const quitAt = shellMain.indexOf('app.quit()', flushAt)
    expect(quitAt).toBeGreaterThan(flushAt)
    // flush point 2: will-quit safety net (a quit with zero windows never
    // emits window-all-closed)
    expect(updaterMain).toContain("app.on('will-quit'")
  })

  it('a cancelled dirty-guard disarms the pending install', () => {
    const abortAt = shellMain.indexOf('function abortAppQuit')
    expect(abortAt).toBeGreaterThan(-1)
    expect(shellMain.slice(abortAt)).toContain('cancelPendingInstall()')
  })
})

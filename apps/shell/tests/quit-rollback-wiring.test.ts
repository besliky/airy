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
    const persistAt = shellMain.indexOf('else persistSessionState(decision.skipStaged)')
    const markAt = shellMain.indexOf('entry.closingConfirmed = true')
    expect(persistAt).toBeGreaterThan(-1)
    expect(markAt).toBeGreaterThan(persistAt)
  })
})

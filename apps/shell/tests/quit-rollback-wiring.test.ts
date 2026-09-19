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
})

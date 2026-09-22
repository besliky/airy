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
    const recoveryAt = shellMain.indexOf('if (restoreSession) {')
    expect(rollbackAt).toBeGreaterThan(-1)
    expect(recoveryAt).toBeGreaterThan(rollbackAt)
  })

  it('a failed recovery rewrite is retried and escalated, not ignored (BUG-1311)', () => {
    // the ignored boolean left the quit-time snapshot on disk: the next
    // launch resurrected windows whose close had confirmed before the cancel
    const block = shellMain.slice(
      shellMain.indexOf('function abortAppQuit'),
      shellMain.indexOf('/** which editor'),
    )
    // one retry (a transient failure recovers)...
    expect(block).toContain('!persistSessionState(false) && !persistSessionState(false)')
    // ...then a loud escalation instead of a silent stale snapshot
    expect(block).toContain('console.error')
    expect(block).toContain('recovery rewrite after the aborted quit failed')
  })

  it('the rollback bridge restart rides the toggle serializer (BUG-1312)', () => {
    // a plain startLiveBridge() bypassed the createLiveBridgeToggle chain and
    // could interleave with a concurrent Settings toggle-off
    expect(shellMain).toContain('setLiveBridgeEnabled.runExclusive')
    // the setting is re-checked inside the chain, after any queued toggle
    expect(shellMain).toMatch(
      /runExclusive\(async \(\) => \{\s*\n\s*if \(liveBridgeEnabled\(\)\) await startLiveBridge\(\)/,
    )
  })

  it('the bridge token lives for the process, not per restart (BUG-1313)', () => {
    // shell-bridge.ts owns the cache (it imports Electron, so this source pin
    // is the test): generated once, reused across aborted-quit restarts so
    // the published info file never rotates under connected clients
    const shellBridge = readFileSync(join(here, '../src/main/bridge/shell-bridge.ts'), 'utf8')
    expect(shellBridge).toContain('token: (processToken ??= generateBridgeToken())')
    // server.ts must accept the reuse instead of always generating
    const bridgeServer = readFileSync(join(here, '../src/main/bridge/server.ts'), 'utf8')
    expect(bridgeServer).toContain('options.token ?? generateBridgeToken()')
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

import { describe, expect, it, vi } from 'vitest'

import { createShutdownEffects } from '../src/main/shutdown-effects'

/**
 * Shutdown side-effect rollback (src/main/shutdown-effects.ts): before-quit
 * arms the editor modules' shutdown behavior, and a cancelled quit must
 * unwind it (BUG-1216: the sheets no-prompt flag and the stopped live bridge
 * used to stay in shutdown state forever after one aborted quit — later
 * dirty-tab closes went through silently, without the Save prompt).
 */

interface Deferred {
  promise: Promise<void>
  resolve(): void
}

function deferred(): Deferred {
  let resolve!: () => void
  const promise = new Promise<void>((r) => {
    resolve = r
  })
  return { promise, resolve }
}

function makeHandlers() {
  const stop = deferred()
  return {
    stop,
    setSheetsShuttingDown: vi.fn(),
    stopSidecar: vi.fn(),
    disposePdfWorkers: vi.fn(),
    bridgeEnabled: vi.fn(() => true),
    // the shutdown stop parks on a deferred: tests decide when it settles
    stopBridge: vi.fn(() => stop.promise),
    startBridge: vi.fn(() => Promise.resolve()),
    log: vi.fn(),
  }
}

/** settle pending microtask chains (stop → catch → restart → catch) */
async function flush(): Promise<void> {
  for (let i = 0; i < 10; i++) await Promise.resolve()
}

describe('createShutdownEffects', () => {
  it('arm() arms the shutdown effects across the editor modules', async () => {
    const h = makeHandlers()
    createShutdownEffects(h).arm()
    await flush()
    expect(h.setSheetsShuttingDown).toHaveBeenCalledWith(true)
    expect(h.stopSidecar).toHaveBeenCalledTimes(1)
    expect(h.disposePdfWorkers).toHaveBeenCalledTimes(1)
    expect(h.stopBridge).toHaveBeenCalledTimes(1)
  })

  it('rollback() brings the sheets close prompt back (BUG-1216)', async () => {
    const h = makeHandlers()
    const effects = createShutdownEffects(h)
    effects.arm()
    h.stop.resolve()
    effects.rollback()
    await flush()
    expect(h.setSheetsShuttingDown).toHaveBeenLastCalledWith(false)
  })

  it('rollback() restarts the bridge only after the shutdown stop settled', async () => {
    const h = makeHandlers()
    const effects = createShutdownEffects(h)
    effects.arm()
    effects.rollback()
    // the stop has not settled yet: no restart may bind the socket, or the
    // dying server's token/socket cleanup would delete the new server's files
    await flush()
    expect(h.startBridge).not.toHaveBeenCalled()
    h.stop.resolve()
    await flush()
    expect(h.startBridge).toHaveBeenCalledTimes(1)
  })

  it('rollback() leaves the bridge down when its setting is off', async () => {
    const h = makeHandlers()
    h.bridgeEnabled.mockReturnValue(false)
    const effects = createShutdownEffects(h)
    effects.arm()
    h.stop.resolve()
    effects.rollback()
    await flush()
    expect(h.setSheetsShuttingDown).toHaveBeenCalledWith(false)
    expect(h.startBridge).not.toHaveBeenCalled()
  })

  it('a failed shutdown stop still allows the restart', async () => {
    const h = makeHandlers()
    h.stopBridge = vi.fn(() => Promise.reject(new Error('socket busy')))
    const effects = createShutdownEffects(h)
    effects.arm()
    effects.rollback()
    await flush()
    expect(h.startBridge).toHaveBeenCalledTimes(1)
    expect(h.log).toHaveBeenCalledWith(
      '[shell] bridge stop during shutdown failed:',
      expect.any(Error),
    )
  })

  it('a failed restart is logged, not thrown', async () => {
    const h = makeHandlers()
    h.startBridge = vi.fn(() => Promise.reject(new Error('port taken')))
    const effects = createShutdownEffects(h)
    effects.arm()
    h.stop.resolve()
    effects.rollback()
    await flush()
    expect(h.log).toHaveBeenCalledWith(
      '[shell] bridge restart after an aborted quit failed:',
      expect.any(Error),
    )
  })

  it('a newer arm() cancels a restart queued before it', async () => {
    const h = makeHandlers()
    const effects = createShutdownEffects(h)
    effects.arm() // quit #1 stops the bridge
    effects.rollback() // the quit is aborted: restart queued behind the stop
    effects.arm() // the user re-quits before the stop settled
    h.stop.resolve()
    await flush()
    // the restart belonged to the aborted quit #1; quit #2 owns the bridge
    // now and must keep it down
    expect(h.startBridge).not.toHaveBeenCalled()
    expect(h.setSheetsShuttingDown).toHaveBeenLastCalledWith(true)
  })
})

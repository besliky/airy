/** Live-bridge toggle sequencing (src/main/live-bridge-toggle.ts): the server
 *  transitions before the setting is persisted, failures reject to the caller
 *  without touching the stored value, and concurrent toggles serialize. */
import { describe, expect, it } from 'vitest'

import { createLiveBridgeToggle } from '../src/main/live-bridge-toggle'

interface Call {
  op: 'start' | 'stop' | 'persist'
  on?: boolean
}

function makeDeps(
  overrides: Partial<{ start: () => Promise<void>; stop: () => Promise<void> }> = {},
) {
  const calls: Call[] = []
  let stored: boolean | null = null
  const deps = {
    isEnabled: () => stored ?? true,
    toggleAllowed: () => true,
    persist: (on: boolean) => {
      calls.push({ op: 'persist', on })
      stored = on
    },
    // overrides are wrapped so every attempt is recorded, then their outcome
    // (resolve or reject) is replayed
    start: async () => {
      calls.push({ op: 'start' })
      await overrides.start?.()
    },
    stop: async () => {
      calls.push({ op: 'stop' })
      await overrides.stop?.()
    },
  }
  return { deps, calls, storedRef: () => stored }
}

describe('createLiveBridgeToggle', () => {
  it('transitions the server first and persists the resulting state afterwards', async () => {
    const { deps, calls } = makeDeps()
    const toggle = createLiveBridgeToggle(deps)
    await toggle(false)
    expect(calls).toEqual([{ op: 'stop' }, { op: 'persist', on: false }])
    await toggle(true)
    expect(calls).toEqual([
      { op: 'stop' },
      { op: 'persist', on: false },
      { op: 'start' },
      { op: 'persist', on: true },
    ])
  })

  it('a failing start rejects to the caller and persists nothing (stored value survives)', async () => {
    const { deps, calls, storedRef } = makeDeps({
      start: () => Promise.reject(new Error('EADDRINUSE')),
    })
    const toggle = createLiveBridgeToggle(deps)
    await expect(toggle(true)).rejects.toThrow('EADDRINUSE')
    expect(calls).toEqual([{ op: 'start' }]) // no persist call at all
    expect(storedRef()).toBe(null)
  })

  it('a failing stop also skips the persist and rejects', async () => {
    const { deps, calls } = makeDeps({ stop: () => Promise.reject(new Error('busy')) })
    const toggle = createLiveBridgeToggle(deps)
    await expect(toggle(false)).rejects.toThrow('busy')
    expect(calls).toEqual([{ op: 'stop' }])
  })

  it('serializes two rapid toggles: no interleaving of start/stop/persist', async () => {
    const calls: Call[] = []
    let releaseStart: (() => void) | null = null
    const deps = {
      isEnabled: () => true,
      toggleAllowed: () => true,
      persist: (on: boolean) => calls.push({ op: 'persist', on }),
      start: () => {
        calls.push({ op: 'start' })
        return new Promise<void>((resolve) => {
          releaseStart = resolve
        })
      },
      stop: () => {
        calls.push({ op: 'stop' })
        return Promise.resolve()
      },
    }
    const toggle = createLiveBridgeToggle(deps)
    const first = toggle(true) // hangs in start until released
    const second = toggle(false) // must wait for the first to settle
    await Promise.resolve()
    await Promise.resolve()
    // the second toggle has not touched anything while the first is in flight
    expect(calls).toEqual([{ op: 'start' }])
    releaseStart!()
    await first
    await second
    expect(calls).toEqual([
      { op: 'start' },
      { op: 'persist', on: true },
      { op: 'stop' },
      { op: 'persist', on: false },
    ])
  })

  it('a failed run does not poison the chain: the next toggle still runs', async () => {
    let fail = true
    const { deps, calls } = makeDeps({
      // the makeDeps wrapper records the attempt; the override only outcomes
      start: () => (fail ? Promise.reject(new Error('nope')) : Promise.resolve()),
    })
    const toggle = createLiveBridgeToggle(deps)
    await expect(toggle(true)).rejects.toThrow('nope')
    fail = false
    await expect(toggle(true)).resolves.toBe(true)
    expect(calls).toEqual([{ op: 'start' }, { op: 'start' }, { op: 'persist', on: true }])
  })

  it('is a visible no-op when the env override disallows toggling', async () => {
    const { deps, calls } = makeDeps()
    const disallowed = { ...deps, toggleAllowed: () => false, isEnabled: () => false }
    const toggle = createLiveBridgeToggle(disallowed)
    await expect(toggle(true)).resolves.toBe(false)
    expect(calls).toEqual([])
  })

  it('runExclusive serializes against a pending toggle without persisting (BUG-1312)', async () => {
    // the quit rollback's bridge restart must share the user-toggle order: a
    // plain startLiveBridge() could interleave with a concurrent toggle-off
    const calls: Call[] = []
    let releaseStop: (() => void) | null = null
    const deps = {
      isEnabled: () => true,
      toggleAllowed: () => true,
      persist: (on: boolean) => calls.push({ op: 'persist', on }),
      start: () => {
        calls.push({ op: 'start' })
        return Promise.resolve()
      },
      stop: () => {
        calls.push({ op: 'stop' })
        return new Promise<void>((resolve) => {
          releaseStop = resolve
        })
      },
    }
    const toggle = createLiveBridgeToggle(deps)
    const userOff = toggle(false) // hangs in stop until released
    let restartRan = false
    const restart = toggle.runExclusive(async () => {
      calls.push({ op: 'start' }) // the rollback restart issues its own start
      restartRan = true
    })
    await Promise.resolve()
    await Promise.resolve()
    // the restart has not jumped the in-flight toggle's stop
    expect(restartRan).toBe(false)
    expect(calls).toEqual([{ op: 'stop' }])
    releaseStop!()
    await userOff
    await restart
    // the restart ran after the stop and persisted nothing (the rollback must
    // not overwrite the setting a concurrent toggle just wrote)
    expect(restartRan).toBe(true)
    expect(calls).toEqual([{ op: 'stop' }, { op: 'persist', on: false }, { op: 'start' }])
  })

  it('a failing runExclusive rejects to its caller and keeps the chain alive', async () => {
    const { deps } = makeDeps()
    const toggle = createLiveBridgeToggle(deps)
    await expect(
      toggle.runExclusive(async () => Promise.reject(new Error('boom'))),
    ).rejects.toThrow('boom')
    // the next toggle still runs
    await expect(toggle(true)).resolves.toBe(true)
  })
})

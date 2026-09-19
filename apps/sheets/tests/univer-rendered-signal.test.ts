// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { IRenderManagerService } from '@univerjs/engine-render'

import type { UniverRuntime } from '../src/renderer/univer-state'
import { rearmUniverRenderedSignal } from '../src/renderer/univer-sync'

/**
 * TEST-1101: `data-univer-rendered` means "the CURRENT session is rendered",
 * not "some session once rendered" — Univer reaches Rendered once per
 * runtime, so a session swap must retire the flag synchronously and re-arm
 * it only from the replacement unit's mounted canvas. The e2e helper waits
 * on the attribute; these tests pin the rearm contract itself (remove-then-
 * set order, and the epoch guard against superseded swaps) against a stub
 * runtime, so a regression here cannot hide behind e2e timing slack.
 */

/** a runtime stub whose unit's canvas is (or is not) mounted yet */
function stubRuntime(mountedUnitIds: Set<string>): UniverRuntime {
  const canvas = document.createElement('canvas')
  document.body.append(canvas)
  const renderManager = {
    getRenderById: (unitId: string) =>
      mountedUnitIds.has(unitId)
        ? { mainComponent: {}, engine: { getCanvasElement: () => canvas } }
        : null,
  }
  const injector = {
    get: (token: unknown) => (token === IRenderManagerService ? renderManager : {}),
  }
  return { univer: { __getInjector: () => injector } } as unknown as UniverRuntime
}

const flag = (): string | null => document.documentElement.getAttribute('data-univer-rendered')

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  document.documentElement.removeAttribute('data-univer-rendered')
  document.body.replaceChildren()
  vi.useRealTimers()
})

describe('data-univer-rendered rearm on session swap (TEST-1101)', () => {
  it('drops the flag synchronously and restores it once the new canvas mounts', async () => {
    document.documentElement.setAttribute('data-univer-rendered', 'true')
    const runtime = stubRuntime(new Set(['file-new']))

    rearmUniverRenderedSignal(runtime, 'file-new')

    // the swap retires the previous session's signal immediately — an e2e
    // driver that starts waiting right here must not see the stale session
    expect(flag()).toBeNull()

    // the first poll (100 ms) sees the mounted canvas and defers the set to
    // the next frame, so the first paint of the new session is what lands
    await vi.advanceTimersByTimeAsync(100)
    expect(flag()).toBeNull()
    await vi.advanceTimersByTimeAsync(20)
    expect(flag()).toBe('true')
  })

  it('keeps the flag down while the canvas is unmounted, and a superseded swap never resurrects it', async () => {
    document.documentElement.setAttribute('data-univer-rendered', 'true')
    const runtimeA = stubRuntime(new Set(['file-a']))
    const bMounted = new Set<string>()
    const runtimeB = stubRuntime(bMounted)

    rearmUniverRenderedSignal(runtimeA, 'file-a')
    // session B supersedes A before A's poll fires; B's canvas is not up yet
    rearmUniverRenderedSignal(runtimeB, 'file-b')
    expect(flag()).toBeNull()

    // A's poll runs here and would set the attribute — the epoch guard must
    // drop it instead; B's poll finds no canvas and keeps polling
    await vi.advanceTimersByTimeAsync(100)
    expect(flag()).toBeNull()
    await vi.advanceTimersByTimeAsync(20)
    expect(flag()).toBeNull()
    await vi.advanceTimersByTimeAsync(200)
    expect(flag()).toBeNull()

    // B's canvas finally mounts (the normal createWorkbook latency): only
    // now may the attribute come back, from B's own poll
    bMounted.add('file-b')
    await vi.advanceTimersByTimeAsync(100)
    await vi.advanceTimersByTimeAsync(20)
    expect(flag()).toBe('true')
  })
})

import { describe, expect, it, vi } from 'vitest'

import { createWindowCloseGuard, type GuardedCloseTab } from '../src/main/window-close-guard'

/**
 * Window close guard (src/main/window-close-guard.ts): one walk of a window's
 * dirty tabs through their prompts per close decision (BUG-1217: a second
 * close event during a pending prompt used to start a parallel walk — double
 * dialogs, clashing Cancel/Save answers, the window closing right after a
 * Cancel). Electron, the TabManager, and the per-editor prompt functions are
 * injected; only the guard's own re-entrancy rules are asserted.
 */

interface TestTab extends GuardedCloseTab {
  kind: 'sheets' | 'pdf' | 'docs'
}

function makeTab(id: string, kind: TestTab['kind'] = 'sheets'): TestTab {
  return { id, kind, webContents: { id: Number(id) } as never }
}

function makeGuard(tabs: TestTab[], requestClose: (tab: TestTab) => Promise<boolean>) {
  const deps = {
    dirtyTabs: vi.fn(() => tabs),
    requestClose: vi.fn(requestClose),
    finishClose: vi.fn(),
    closeWindow: vi.fn(),
    isWindowAlive: vi.fn(() => true),
    abortQuit: vi.fn(),
    logFailure: vi.fn(),
  }
  return { deps, guard: createWindowCloseGuard(deps) }
}

function closeEvent(): { preventDefault: () => void; preventDefaultSpy: ReturnType<typeof vi.fn> } {
  const preventDefaultSpy = vi.fn()
  return { preventDefault: preventDefaultSpy, preventDefaultSpy }
}

/** settle the guard cycle's microtasks */
async function flush(): Promise<void> {
  for (let i = 0; i < 10; i++) await Promise.resolve()
}

describe('createWindowCloseGuard', () => {
  it('a clean window closes without a guard cycle or preventDefault', () => {
    const { deps, guard } = makeGuard([], async () => true)
    const event = closeEvent()
    guard(event)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(event.preventDefaultSpy).not.toHaveBeenCalled()
    expect(deps.requestClose).not.toHaveBeenCalled()
  })

  it('all tabs confirm: bookkeeping once, then the window closes for real', async () => {
    const { deps, guard } = makeGuard([makeTab('1'), makeTab('2', 'pdf')], async () => true)
    guard(closeEvent())
    await flush()
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).toHaveBeenCalledTimes(1)
    expect(deps.abortQuit).not.toHaveBeenCalled()
    // the re-fired close event passes through without another cycle
    guard(closeEvent())
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
  })

  it('a cancelled prompt aborts the quit and leaves the window open', async () => {
    const { deps, guard } = makeGuard([makeTab('1')], async () => false)
    guard(closeEvent())
    await flush()
    expect(deps.abortQuit).toHaveBeenCalledTimes(1)
    expect(deps.finishClose).not.toHaveBeenCalled()
    expect(deps.closeWindow).not.toHaveBeenCalled()
  })

  it('a repeat close event while the prompt is pending runs NO second cycle (BUG-1217)', async () => {
    let answer!: (ok: boolean) => void
    const { deps, guard } = makeGuard(
      [makeTab('1')],
      () => new Promise<boolean>((resolve) => (answer = resolve)),
    )
    // X button opens the prompt; Cmd+Q fires a second close on every window
    guard(closeEvent())
    const repeat = closeEvent()
    guard(repeat)
    expect(repeat.preventDefaultSpy).toHaveBeenCalledTimes(1)
    expect(deps.requestClose).toHaveBeenCalledTimes(1) // ONE dialog for the tab
    expect(deps.finishClose).not.toHaveBeenCalled()
    // the user cancels the one open dialog: the window must stay open and
    // the quit abort exactly once
    answer(false)
    await flush()
    expect(deps.abortQuit).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).not.toHaveBeenCalled()
  })

  it('the repeat event is swallowed even when it would confirm', async () => {
    let answer!: (ok: boolean) => void
    const { deps, guard } = makeGuard(
      [makeTab('1')],
      () => new Promise<boolean>((resolve) => (answer = resolve)),
    )
    guard(closeEvent())
    guard(closeEvent()) // the repeat is dropped, not queued
    answer(true)
    await flush()
    expect(deps.requestClose).toHaveBeenCalledTimes(1)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).toHaveBeenCalledTimes(1)
  })

  it('a cancelled cycle releases the latch: a later close starts a fresh cycle', async () => {
    let answer!: (ok: boolean) => void
    const { deps, guard } = makeGuard(
      [makeTab('1')],
      () => new Promise<boolean>((resolve) => (answer = resolve)),
    )
    guard(closeEvent())
    answer(false)
    await flush()
    // the user tries again: a new guard cycle must run
    guard(closeEvent())
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
    answer(true)
    await flush()
    expect(deps.closeWindow).toHaveBeenCalledTimes(1)
  })

  it('a destroyed window is not closed again after the cycle confirms', async () => {
    const { deps, guard } = makeGuard([makeTab('1')], async () => true)
    deps.isWindowAlive.mockReturnValue(false)
    guard(closeEvent())
    await flush()
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).not.toHaveBeenCalled()
  })

  it('a throwing prompt is logged and releases the latch, not the window', async () => {
    const { deps, guard } = makeGuard([makeTab('1')], async () => {
      throw new Error('renderer died mid-prompt')
    })
    guard(closeEvent())
    await flush()
    expect(deps.logFailure).toHaveBeenCalledWith(expect.any(Error))
    expect(deps.finishClose).not.toHaveBeenCalled()
    expect(deps.closeWindow).not.toHaveBeenCalled()
    // the failed cycle is gone: a new close runs a fresh cycle instead of
    // hitting a stranded latch
    guard(closeEvent())
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
  })
})

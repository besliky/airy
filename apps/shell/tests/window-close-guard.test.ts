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

  it('a throwing prompt round-trip unwinds an in-flight quit like a Cancel (BUG-1310)', async () => {
    // the failure path is not a user decision, but its aftermath must match
    // the Cancel path: a quit that was armed when the round-trip blew up
    // (contents.send on a destroyed webContents) used to stay armed forever
    // — no window closed, sheets kept the quit-time no-prompt mode, the
    // bridge stayed down
    const { deps, guard } = makeGuard([makeTab('1'), makeTab('2', 'pdf')], async (tab) => {
      if (tab.id === '2') throw new Error('webContents destroyed mid-send')
      return true
    })
    guard(closeEvent())
    await flush()
    expect(deps.logFailure).toHaveBeenCalledTimes(1)
    // the quit unwinds once, the failed window neither closes nor finishes
    expect(deps.abortQuit).toHaveBeenCalledTimes(1)
    expect(deps.finishClose).not.toHaveBeenCalled()
    expect(deps.closeWindow).not.toHaveBeenCalled()
    // and the latch is free: the next close starts and completes a fresh
    // cycle over both tabs
    guard(closeEvent())
    await flush()
    expect(deps.requestClose).toHaveBeenCalledTimes(4)
  })
})

describe('dirty-set recalc before the final close (BUG-410)', () => {
  /** a guard whose dirty set and answers are driven by the test */
  function makeRecalcGuard(
    dirtyTabs: () => TestTab[],
    requestClose: (tab: TestTab) => Promise<boolean> = async () => true,
  ) {
    const deps = {
      dirtyTabs: vi.fn(dirtyTabs),
      requestClose: vi.fn(requestClose),
      finishClose: vi.fn(),
      closeWindow: vi.fn(),
      isWindowAlive: vi.fn(() => true),
      abortQuit: vi.fn(),
      logFailure: vi.fn(),
    }
    return { deps, guard: createWindowCloseGuard(deps) }
  }

  it('a tab that became dirty mid-cycle is prompted before the close confirms', async () => {
    // tab 3 appears while the walk over tabs 1-2 is still prompting (a
    // background edit, an AI flow opening a document) — the confirmed close
    // used to swallow it silently
    const tab3 = makeTab('3', 'pdf')
    const dirty = vi.fn<() => TestTab[]>().mockReturnValueOnce([makeTab('1'), makeTab('2')])
    dirty.mockReturnValueOnce([makeTab('1'), makeTab('2'), tab3])
    dirty.mockReturnValue([makeTab('1'), makeTab('2'), tab3])
    const { deps, guard } = makeRecalcGuard(dirty)
    guard(closeEvent())
    await flush()
    expect(deps.requestClose).toHaveBeenCalledTimes(3)
    expect(deps.requestClose).toHaveBeenLastCalledWith(tab3)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).toHaveBeenCalledTimes(1)
  })

  it('tabs the cycle already walked are not re-prompted by the recalc pass', async () => {
    // "Don't save" leaves a tab dirty by design: the stable dirty list must
    // not make the second pass ask the same user the same question again
    const { deps, guard } = makeRecalcGuard(() => [makeTab('1'), makeTab('2')])
    guard(closeEvent())
    await flush()
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).toHaveBeenCalledTimes(1)
  })

  it('the recalc is bounded at two passes when tabs keep dirtying', async () => {
    const grows = [
      [makeTab('1')],
      [makeTab('1'), makeTab('2', 'pdf')],
      [makeTab('1'), makeTab('2', 'pdf'), makeTab('3', 'docs')],
      [makeTab('1'), makeTab('2', 'pdf'), makeTab('3', 'docs'), makeTab('4')],
    ]
    let call = 0
    const { deps, guard } = makeRecalcGuard(() => grows[Math.min(call++, grows.length - 1)]!)
    guard(closeEvent())
    await flush()
    // pass 1 walked tab 1, pass 2 walked tab 2 — then the close goes through
    // instead of chasing the ever-growing set forever
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
    expect(deps.finishClose).toHaveBeenCalledTimes(1)
    expect(deps.closeWindow).toHaveBeenCalledTimes(1)
  })

  it('a Cancel on a mid-cycle-dirtied tab aborts the quit like any other', async () => {
    const tab2 = makeTab('2', 'pdf')
    const dirty = vi.fn<() => TestTab[]>().mockReturnValueOnce([makeTab('1')])
    dirty.mockReturnValue([makeTab('1'), tab2])
    const { deps, guard } = makeRecalcGuard(dirty, async (tab) => tab.id !== '2')
    guard(closeEvent())
    await flush()
    expect(deps.requestClose).toHaveBeenCalledTimes(2)
    expect(deps.abortQuit).toHaveBeenCalledTimes(1)
    expect(deps.finishClose).not.toHaveBeenCalled()
    expect(deps.closeWindow).not.toHaveBeenCalled()
  })
})

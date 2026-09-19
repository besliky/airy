import { describe, expect, it, vi } from 'vitest'

/**
 * Shell window registry (src/main/shell-windows.ts): which windows are live,
 * which one Home IPC senders belong to, and how focus resolves while the
 * user moves between shell windows, editor windows, and dialogs. Electron
 * and TabManager are opaque here — only the registry's own routing rules
 * are asserted.
 */

vi.mock('electron', () => ({ BrowserWindow: class {} }))

import { ShellWindowRegistry, type ShellWindowEntry } from '../src/main/shell-windows'

function makeEntry(id: number, homeWebContentsId: number): ShellWindowEntry {
  return {
    // unique object identity per fake window — what the registry keys on
    win: { id } as never,
    manager: {} as never,
    homeWebContentsId,
    homeRendererCrashed: false,
    primary: id === 1,
  }
}

describe('ShellWindowRegistry', () => {
  it('lists entries in creation order (the session persistence order)', () => {
    const registry = new ShellWindowRegistry()
    const first = makeEntry(1, 10)
    const second = makeEntry(2, 20)
    registry.add(first)
    registry.add(second)
    expect(registry.list()).toEqual([first, second])
  })

  it('the first added window stands in as focused until a focus event arrives', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntry(1, 10)
    registry.add(primary)
    expect(registry.focused()).toBe(primary)
  })

  it('notifyFocused moves the resolution target and survives non-shell windows', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntry(1, 10)
    const secondary = makeEntry(2, 20)
    registry.add(primary)
    registry.add(secondary)

    // focus moving to a non-shell window (editor/password dialog) is not
    // reported; the last focused shell window keeps resolving
    expect(registry.focused()).toBe(primary)
    registry.notifyFocused(secondary.win)
    expect(registry.focused()).toBe(secondary)
  })

  it('resolves Home IPC senders to their own window', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntry(1, 10)
    const secondary = makeEntry(2, 20)
    registry.add(primary)
    registry.add(secondary)

    expect(registry.forHomeWebContents(10)).toBe(primary)
    expect(registry.forHomeWebContents(20)).toBe(secondary)
    expect(registry.forHomeWebContents(999)).toBeUndefined()
  })

  it('removing a window drops it and passes focus to the first survivor', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntry(1, 10)
    const secondary = makeEntry(2, 20)
    registry.add(primary)
    registry.add(secondary)
    registry.notifyFocused(secondary)

    expect(registry.remove(secondary.win)).toBe(secondary)
    expect(registry.list()).toEqual([primary])
    expect(registry.focused()).toBe(primary)
    expect(registry.forHomeWebContents(20)).toBeUndefined()
    // removing an unknown window is a no-op
    expect(registry.remove({} as never)).toBeUndefined()
  })

  it('an empty registry resolves nothing', () => {
    const registry = new ShellWindowRegistry()
    expect(registry.list()).toEqual([])
    expect(registry.focused()).toBeUndefined()
  })

  /** entry whose fake manager hosts one editor view webContents (wcId) */
  function makeEntryWithTab(id: number, homeWebContentsId: number, wcId: number) {
    const entry = makeEntry(id, homeWebContentsId)
    const manager = { tabIdForWebContents: (wc: number) => (wc === wcId ? `t${id}` : undefined) }
    return Object.assign(entry, { manager }) as ShellWindowEntry & { manager: typeof manager }
  }

  it('managerForWebContents finds the strip hosting the editor view', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntryWithTab(1, 10, 101)
    const secondary = makeEntryWithTab(2, 20, 202)
    registry.add(primary)
    registry.add(secondary)

    expect(registry.managerForWebContents(101)).toBe(primary.manager)
    expect(registry.managerForWebContents(202)).toBe(secondary.manager)
    expect(registry.managerForWebContents(999)).toBeNull()
  })

  it('managerForSender answers in the sender window, not the focused one (BUG-1107)', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntryWithTab(1, 10, 101)
    const secondary = makeEntryWithTab(2, 20, 202)
    registry.add(primary)
    registry.add(secondary)
    // window 1 holds focus while a background tab of window 2 asks
    registry.notifyFocused(primary.win)

    // a background tab of the UNFOCUSED window gets its own manager...
    expect(registry.managerForSender(202)).toBe(secondary.manager)
    // ...an unknown sender falls back to the focused window...
    expect(registry.managerForSender(999)).toBe(primary.manager)
    // ...and menu-driven calls (no sender) also resolve by focus
    expect(registry.managerForSender(undefined)).toBe(primary.manager)
  })
})

describe('persistableEntries (session write source)', () => {
  it('serializes every registered window in creation order', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntry(1, 10)
    const secondary = makeEntry(2, 20)
    registry.add(primary)
    registry.add(secondary)
    expect(registry.persistableEntries()).toEqual([primary, secondary])
  })

  it('leaves out the window its own ordinary close excludes', () => {
    const registry = new ShellWindowRegistry()
    const primary = makeEntry(1, 10)
    const secondary = makeEntry(2, 20)
    registry.add(primary)
    registry.add(secondary)
    expect(registry.persistableEntries(secondary)).toEqual([primary])
  })

  it('leaves out a confirmed-closing window until its closed event lands (BUG-1218)', () => {
    const registry = new ShellWindowRegistry()
    const confirming = makeEntry(1, 10)
    const surviving = makeEntry(2, 20)
    const lateGuard = makeEntry(3, 30)
    registry.add(confirming)
    registry.add(surviving)
    registry.add(lateGuard)
    // finishWindowClose marked the closer; its asynchronous 'closed' event
    // has not fired yet, so it is still registered and listed...
    confirming.closingConfirmed = true
    expect(registry.list()).toEqual([confirming, surviving, lateGuard])
    // ...but no session write may resurrect it: the recovery rewrite after
    // another window cancels the quit, and any ordinary close that lands in
    // the same gap, serialize the survivors only
    expect(registry.persistableEntries()).toEqual([surviving, lateGuard])
    expect(registry.persistableEntries(lateGuard)).toEqual([surviving])
    // once 'closed' arrives the registry forgets it entirely, flag or not
    registry.remove(confirming.win)
    expect(registry.persistableEntries()).toEqual([surviving, lateGuard])
  })
})

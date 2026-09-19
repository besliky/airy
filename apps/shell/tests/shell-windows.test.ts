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
})

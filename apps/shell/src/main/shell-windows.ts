import type { BrowserWindow } from 'electron'
import type { TabManager } from './tab-manager'

/**
 * One open shell window: the BrowserWindow, the TabManager owning its strip,
 * and the window's own Home renderer identity. The shell supports several of
 * these at once ("Move tab to new window"); the first one created is the
 * primary (it owns the persisted geometry and receives OS-level file opens
 * when nothing is focused).
 */
export interface ShellWindowEntry {
  win: BrowserWindow
  manager: TabManager
  /** this window's own Home renderer webContents id (home:* sender check) */
  homeWebContentsId: number
  /** Home renderer crashed and awaits its Reload decision (dedupe guard) */
  homeRendererCrashed: boolean
  primary: boolean
}

/**
 * Live shell windows in creation order — the order session persistence uses.
 * Pure bookkeeping, no Electron calls beyond the passed-in window objects, so
 * the routing decisions are unit-testable with fake windows.
 */
export class ShellWindowRegistry {
  private readonly entries: ShellWindowEntry[] = []
  private lastFocused: ShellWindowEntry | null = null

  add(entry: ShellWindowEntry): void {
    this.entries.push(entry)
    // before any focus event arrives, the first window stands in as the
    // focused one (later windows take focus via their own focus event)
    if (!this.lastFocused) this.lastFocused = entry
  }

  /** drop a destroyed window; focus falls back to the first survivor */
  remove(win: BrowserWindow): ShellWindowEntry | undefined {
    const idx = this.entries.findIndex((e) => e.win === win)
    if (idx < 0) return undefined
    const [removed] = this.entries.splice(idx, 1)
    if (this.lastFocused === removed) this.lastFocused = this.entries[0] ?? null
    return removed
  }

  list(): readonly ShellWindowEntry[] {
    return this.entries
  }

  forWindow(win: BrowserWindow): ShellWindowEntry | undefined {
    return this.entries.find((e) => e.win === win)
  }

  /** the entry whose Home renderer (shell window's own webContents) this is */
  forHomeWebContents(webContentsId: number): ShellWindowEntry | undefined {
    return this.entries.find((e) => e.homeWebContentsId === webContentsId)
  }

  /** the manager whose strip hosts this editor view (a tab lives in exactly
   *  one window); null when no open tab owns the webContents */
  managerForWebContents(webContentsId: number): TabManager | null {
    for (const entry of this.entries) {
      if (entry.manager.tabIdForWebContents(webContentsId) !== undefined) return entry.manager
    }
    return null
  }

  /** Routing target for a hook call that can name its sender: the window
   *  whose strip hosts the sender's tab, falling back to the focused one
   *  when no sender came along (menu-initiated calls) or its tab is already
   *  gone. Resolving by focus alone teleported background-tab results (AI
   *  document creation, exports) into whichever window happened to be
   *  focused (BUG-1107). */
  managerForSender(senderWcId: number | undefined): TabManager | null {
    const resolved = senderWcId === undefined ? null : this.managerForWebContents(senderWcId)
    return resolved ?? this.focused()?.manager ?? null
  }

  /** record a focus event so focused() survives focus moving to non-shell windows */
  notifyFocused(win: BrowserWindow): void {
    const entry = this.forWindow(win)
    if (entry) this.lastFocused = entry
  }

  /** the entry user actions resolve against: the last focused shell window, else the first */
  focused(): ShellWindowEntry | undefined {
    return this.lastFocused ?? this.entries[0]
  }
}

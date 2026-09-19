import type { WebContents } from 'electron'

/**
 * A shell window's close guard, extracted from the window wiring so the
 * re-entrancy rules are unit-testable with fake prompts.
 *
 * The guard walks the window's dirty tabs through their save/don't-save/
 * cancel prompts; any cancel aborts the close (and with it an in-flight
 * app quit — the caller's abortQuit). A second close event while a cycle is
 * still awaiting a prompt (the X button, then Cmd+Q firing close on every
 * window) used to start a SECOND parallel cycle over the same tabs: double
 * dialogs for one tab, clashing Cancel/Save answers, the window closing
 * right after the user had just cancelled (BUG-1217). The pending latch
 * makes the running cycle own the decision — the repeat event is merely
 * prevented from closing the window and dropped.
 */
export interface GuardedCloseTab {
  id: string
  webContents: WebContents
}

export interface WindowCloseGuardDeps<Tab extends GuardedCloseTab> {
  /** dirty (or dirtiness-unknown docs) tabs in walk order; empty = clean */
  dirtyTabs(): readonly Tab[]
  /** one tab's prompt + save round-trip; false = the user cancelled */
  requestClose(tab: Tab): Promise<boolean>
  /** bookkeeping once the close is really going through (clean and
   *  confirmed paths alike: staged cleanup + session persist) */
  finishClose(): void
  /** re-fire the window close the confirmed flag then lets through */
  closeWindow(): void
  /** false once the BrowserWindow is destroyed */
  isWindowAlive(): boolean
  /** a dirty-guard Cancel: unwind an in-flight quit (BUG-1104) */
  abortQuit(): void
  /** a prompt/save round-trip threw: log-only (the cycle is abandoned and
   *  the latch released — the pre-extraction code surfaced the same failure
   *  as an unhandled rejection the process handler logged) */
  logFailure(err: unknown): void
}

export function createWindowCloseGuard<Tab extends GuardedCloseTab>(
  deps: WindowCloseGuardDeps<Tab>,
): (event: { preventDefault(): void }) => void {
  let closeConfirmed = false
  let guardPending = false
  return function onWindowClose(event: { preventDefault(): void }): void {
    if (closeConfirmed) return
    const dirty = deps.dirtyTabs()
    if (dirty.length === 0) {
      // clean window: no guard cycle, the close simply goes through
      deps.finishClose()
      return
    }
    // BUG-1217: a repeat close event (Cmd+Q while this window's X-button
    // prompt is still open) must not start a second parallel cycle over the
    // same tabs — the running one owns the decision
    const firstCycle = !guardPending
    if (firstCycle) guardPending = true
    event.preventDefault()
    if (!firstCycle) return
    void (async () => {
      try {
        for (const tab of dirty) {
          if (!(await deps.requestClose(tab))) {
            deps.abortQuit()
            return
          }
        }
        closeConfirmed = true
        deps.finishClose()
        if (deps.isWindowAlive()) deps.closeWindow()
      } catch (err) {
        deps.logFailure(err)
      } finally {
        // never strand the latch: a cancelled or failed cycle releases it so
        // a later close event starts a fresh one; a confirmed close keeps it
        // up (harmless — closeConfirmed short-circuits this handler)
        if (!closeConfirmed) guardPending = false
      }
    })()
  }
}

import type { WebContents } from 'electron'

/**
 * A shell window's close guard, extracted from the window wiring so the
 * re-entrancy rules are unit-testable with fake prompts.
 *
 * The guard walks the window's dirty tabs through their save/don't-save/
 * cancel prompts; any cancel — or a prompt/save round-trip that throws
 * (BUG-1310) — aborts the close and unwinds an in-flight app quit with it
 * (the caller's abortQuit). A second close event while a cycle is
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
  /** a prompt/save round-trip threw: log it — the abandoned cycle then
   *  unwinds an in-flight quit exactly like a Cancel would (BUG-1310) */
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
        // BUG-410: `dirty` was captured when the close event arrived; a tab
        // that became dirty while these prompts ran (a background edit, an AI
        // flow opening a document in this window) used to be closed silently
        // by the confirmed flag. Recompute the set before the final close and
        // walk only the tabs this cycle has not already put through a prompt
        // ("Don't save" leaves its tab dirty by design, so a plain second
        // walk would re-prompt it); two passes bound the loop against a flow
        // that keeps dirtying further tabs while the current ones confirm.
        const walked = new Set(dirty.map((tab) => tab.id))
        for (let pass = 0; pass < 2; pass += 1) {
          const pending = pass === 0 ? dirty : deps.dirtyTabs().filter((tab) => !walked.has(tab.id))
          for (const tab of pending) {
            walked.add(tab.id)
            if (!(await deps.requestClose(tab))) {
              deps.abortQuit()
              return
            }
          }
        }
        closeConfirmed = true
        deps.finishClose()
        if (deps.isWindowAlive()) deps.closeWindow()
      } catch (err) {
        // BUG-1310: the abandoned cycle must unwind an armed quit exactly
        // like a Cancel does (abortQuit is a no-op when no quit is in
        // flight). A thrown prompt/save round-trip during a quit (e.g.
        // contents.send on a destroyed webContents) used to only release
        // the latch: the quit stayed armed with no window ever closing —
        // sheets kept closing dirty tabs with no prompt and the live
        // bridge stayed down until restart (the BUG-1216 symptom via a new
        // path).
        deps.logFailure(err)
        deps.abortQuit()
      } finally {
        // never strand the latch: a cancelled or failed cycle releases it so
        // a later close event starts a fresh one; a confirmed close keeps it
        // up (harmless — closeConfirmed short-circuits this handler)
        if (!closeConfirmed) guardPending = false
      }
    })()
  }
}

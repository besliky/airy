/**
 * Quit-vs-ordinary-close bookkeeping for the shell's window-close path.
 *
 * `before-quit` arms a quit; every shell window then walks its dirty-tab
 * guards, and ANY cancelled guard aborts the whole quit while the other
 * windows keep running. This state machine keeps that unwind explicit and
 * unit-testable: after a cancel the app must behave exactly like one that
 * never started quitting — later ordinary closes take the per-window branch
 * again (BUG-1104: the `quitting` flag used to stay armed forever, so every
 * later window close took the quit branch and never wrote its per-window
 * session, silently resurrecting closed windows after a crash).
 */
export interface WindowClosePersistDecision {
  /** write the session at all — a quit writes ONE snapshot at the first
   *  confirmed window close; later confirmed closes must not shrink it by
   *  re-serializing without the already-closed windows */
  persist: boolean
  /** drop never-saved untitled tabs from the snapshot */
  skipStaged: boolean
  /** leave the closing window out of the snapshot (one of several windows
   *  closing normally — the survivors are re-serialized without it) */
  excludeClosing: boolean
}

export interface QuitFlow {
  readonly quitting: boolean
  /** before-quit: window closes from here on belong to the app-wide quit.
   *  A repeat while the quit is still in flight is a no-op (BUG-1219) —
   *  re-arming happens only after cancel() reset the flags. */
  begin(): void
  /** a dirty-guard Cancel stopped a window close while windows stay open;
   *  returns true when a quit-time session snapshot has already landed and
   *  must be replaced by a fresh serialize of the live windows */
  cancel(): boolean
  /** persist decision for a window close that passed its guards.
   *  liveWindows counts every registered shell window including the closing
   *  one (it runs before the window leaves the registry). */
  closeDecision(liveWindows: number): WindowClosePersistDecision
  /** the quit snapshot write failed after closeDecision already armed
   *  persist-once (BUG-1224): disarm it so the NEXT confirmed close retries
   *  the write instead of skipping it and leaving a stale session that
   *  resurrects windows closed before the failure. */
  markSnapshotWriteFailed(): void
  /** true while a quit is in flight and its one snapshot has not landed —
   *  an armed persist-once, a failed write, or a quit whose windows all
   *  closed before any snapshot could be written. The final will-quit flush
   *  retries the write for the BUG-1307 residual: a failure on the LAST
   *  confirmed close has no successor close to re-arm into. */
  snapshotWritePending(): boolean
}

export function createQuitFlow(): QuitFlow {
  let quitting = false
  let quitSessionPersisted = false
  return {
    get quitting() {
      return quitting
    },
    begin() {
      // a second before-quit while the quit is still in flight (repeat Cmd+Q
      // with a window-modal prompt still open, or a programmatic app.quit()
      // from the updater) must NOT re-arm the snapshot counter (BUG-1219):
      // the first confirmed close already wrote the one quit snapshot, and a
      // re-armed counter would let the next confirmed close overwrite it
      // without the already-closed windows — shrinking a snapshot this
      // module promises never to shrink — and make a later cancel() look
      // like nothing persisted, skipping the recovery rewrite
      if (quitting) return
      quitting = true
      quitSessionPersisted = false
    },
    cancel() {
      const persistedDuringQuit = quitting && quitSessionPersisted
      quitting = false
      quitSessionPersisted = false
      return persistedDuringQuit
    },
    markSnapshotWriteFailed() {
      // no-op outside a quit: an ordinary-close write failure is retried by
      // the debounced save anyway, and there is no persist-once to re-arm
      if (quitting) quitSessionPersisted = false
    },
    snapshotWritePending() {
      return quitting && !quitSessionPersisted
    },
    closeDecision(liveWindows) {
      if (quitting) {
        // app-wide quit: one write at the first confirmed close keeps every
        // still-registered window's file-backed tabs for the next launch.
        // No skipStaged here (BUG-1105): other windows' guards may still be
        // pending, and stripping their never-saved tabs now would lose the
        // crash-restore safety net for documents whose windows are still
        // live — a cancelled guard would leave unsaved untitled work both
        // unrestorable and purged as orphans on the next launch. Staged
        // files of windows that DO confirm are deleted by their own close
        // before this write runs, and restore prunes nonexistent paths
        // anyway.
        if (quitSessionPersisted)
          return { persist: false, skipStaged: false, excludeClosing: false }
        quitSessionPersisted = true
        return { persist: true, skipStaged: false, excludeClosing: false }
      }
      if (liveWindows > 1) {
        // ordinary close of one of several windows: the survivors keep their
        // staged tabs (they are still live); the closing window drops out
        return { persist: true, skipStaged: false, excludeClosing: true }
      }
      // ordinary close of the last window: the old single-window quit
      // semantics — its file-backed tabs survive in the session, staged ones
      // (whose files this close just deleted) do not
      return { persist: true, skipStaged: true, excludeClosing: false }
    },
  }
}

/**
 * Shutdown side effects and their rollback (BUG-1216).
 *
 * `before-quit` arms process-wide shutdown effects: the sheets close guard's
 * no-prompt mode, the sheets sidecar, the live bridge, and the pdf conversion
 * workers. A cancelled dirty-guard then aborts the quit while the app keeps
 * running — those effects used to stay armed forever: after ONE cancelled
 * quit, closing a dirty sheets tab went through silently (no Save / Don't
 * Save / Cancel prompt, no journal save — only the 30 s recovery copy was
 * left), and the live bridge socket stayed down although its setting was on.
 * This coordinator owns the arm/rollback pair in one place: rollback
 * restores the prompt mode and restarts the bridge when its setting says so,
 * sequenced AFTER the shutdown stop so the dying server's cleanup (socket +
 * token file removal) cannot land on the freshly restarted one.
 */
export interface ShutdownEffectHandlers {
  /** sheets close guard no-prompt mode (on while the app is shutting down) */
  setSheetsShuttingDown(on: boolean): void
  /** stop the sheets sidecar process (lazily restarted on next use — no rollback) */
  stopSidecar(): void
  /** kill in-flight pdf->docx conversion workers (one-shot damage — no rollback) */
  disposePdfWorkers(): void
  /** true when the live bridge should be running (Settings; re-read at rollback
   *  so a bridge the user disabled mid-quit stays disabled) */
  bridgeEnabled(): boolean
  /** stop the live bridge (closes the socket and removes the token file) */
  stopBridge(): Promise<unknown> | unknown
  /** start the live bridge again after an aborted quit */
  startBridge(): Promise<unknown> | unknown
  /** best-effort logging for arm/rollback failures (must not throw) */
  log(message: string, err: unknown): void
}

export interface ShutdownEffects {
  /** before-quit: arm the shutdown behavior across the editor modules */
  arm(): void
  /** a cancelled dirty-guard aborted the quit: unwind what arm() did */
  rollback(): void
}

export function createShutdownEffects(h: ShutdownEffectHandlers): ShutdownEffects {
  // the latest stop issued by arm(); a restart must wait for it to settle, or
  // the stopped server's socket/token cleanup would delete the restarted
  // server's files right after it binds
  let bridgeStop: Promise<unknown> = Promise.resolve()
  // bumped by every arm(): a rollback queued BEFORE the latest arm must not
  // start a bridge the newer quit has just stopped
  let generation = 0
  return {
    arm() {
      generation++
      h.setSheetsShuttingDown(true)
      h.stopSidecar()
      h.disposePdfWorkers()
      // swallowing into bridgeStop keeps the shutdown path rejection-free
      // while still surfacing the failure
      bridgeStop = Promise.resolve()
        .then(() => h.stopBridge())
        .catch((err: unknown) => {
          h.log('[shell] bridge stop during shutdown failed:', err)
        })
    },
    rollback() {
      h.setSheetsShuttingDown(false)
      if (!h.bridgeEnabled()) return
      const armGeneration = generation
      void bridgeStop
        .then(() => {
          // a newer before-quit re-armed while this rollback waited — the
          // bridge must stay down for it
          if (armGeneration !== generation) return
          return h.startBridge()
        })
        .catch((err: unknown) => {
          h.log('[shell] bridge restart after an aborted quit failed:', err)
        })
    },
  }
}

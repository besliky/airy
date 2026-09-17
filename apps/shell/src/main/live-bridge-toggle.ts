/**
 * Live-bridge toggle sequencing: bring the server up/down FIRST, persist the
 * preference only after that succeeds, and serialize concurrent toggles.
 *
 * The old order (persist, then start/stop) left `true` on disk with the
 * server down when a start failed, and two rapid toggles could interleave
 * their async start/stop with each other. The sequencer below is pure (all
 * effects injected) so the ordering and failure semantics are unit-testable;
 * apps/shell/src/main/index.ts wires the real bridge and settings store.
 */
export interface LiveBridgeToggleDeps {
  /** effective enabled state (persisted value + env override resolution) */
  isEnabled(): boolean
  /** false when AIRY_DISABLE_BRIDGE=1 pins the bridge: the toggle is a no-op */
  toggleAllowed(): boolean
  persist(on: boolean): void
  start(): Promise<void>
  stop(): Promise<void>
}

export function createLiveBridgeToggle(
  deps: LiveBridgeToggleDeps,
): (on: boolean) => Promise<boolean> {
  // every toggle runs after the previous one settled (success or failure):
  // two rapid flips cannot interleave start/stop
  let chain: Promise<void> = Promise.resolve()
  return (on: boolean): Promise<boolean> => {
    if (!deps.toggleAllowed()) return Promise.resolve(deps.isEnabled())
    const run = chain.then(async () => {
      if (on) await deps.start()
      else await deps.stop()
      // reached only when the transition succeeded: persist the EFFECTIVE
      // resulting state — a failed start throws below and leaves the
      // previous stored value intact
      deps.persist(on)
    })
    // keep the chain alive regardless of this run's outcome; `run` itself
    // rejects so the caller (the IPC result) sees the failure
    chain = run.then(
      () => undefined,
      () => undefined,
    )
    return run.then(() => deps.isEnabled())
  }
}

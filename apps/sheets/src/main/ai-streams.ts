/**
 * Registry of the in-flight AI streams of one sheets tab (BUG-407).
 *
 * Each streamed turn runs under an AbortController owned by the main process
 * (aiStream/aiStreamCancel IPC). The registry ties those controllers to the
 * tab lifecycle: when the tab's webContents is destroyed, `abortAll` cancels
 * every still-running provider request so a closed tab stops accumulating
 * stream chunks instead of draining the provider response into a dead sender.
 */
export class AiStreamRegistry {
  private readonly streams = new Map<string, AbortController>()

  /** Registers the abort controller of one starting stream. */
  start(requestId: string): AbortController {
    const controller = new AbortController()
    this.streams.set(requestId, controller)
    return controller
  }

  /** Drops the entry once a stream settled (done, error, or abort). */
  finish(requestId: string): void {
    this.streams.delete(requestId)
  }

  /** Aborts one in-flight stream (renderer cancel request); unknown ids are a no-op. */
  cancel(requestId: string): void {
    this.streams.get(requestId)?.abort()
  }

  /** Aborts every in-flight stream — tab teardown. Settled streams are already gone. */
  abortAll(): void {
    for (const controller of this.streams.values()) controller.abort()
    this.streams.clear()
  }
}

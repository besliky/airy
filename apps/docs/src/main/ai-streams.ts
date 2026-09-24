import { AiStreamRegistry } from '../../../sheets/src/main/ai-streams'

/**
 * Minimal structural subset of Electron's WebContents that the registry
 * touches, so unit tests can drive the lifecycle with a fake sender.
 */
export interface AiStreamSender {
  readonly id: number
  once(channel: string, listener: () => void): unknown
  removeListener(channel: string, listener: () => void): unknown
}

/**
 * Per-renderer lifecycle for the app-wide ai:stream registry (BUG-1698).
 *
 * The generic ai:* handlers are registered once per process by the shell
 * (docs-main registerAiIpc) and serve every window type; each stream is bound
 * to the `event.sender` webContents that asked for it. A renderer that dies
 * mid-turn (crash, oom, kill -9) never sends ai:stream-cancel, so — the
 * per-tab BUG-407 teardown adapted to per-sender bookkeeping — every stream
 * of a dead sender must be aborted: this class reuses AiStreamRegistry and
 * aborts a sender's whole registry when its webContents reports
 * 'render-process-gone' (hard death: a killed renderer leaves the webContents
 * object alive for a while, so 'destroyed' alone observes it too late) or
 * 'destroyed' (graceful teardown, same semantics as sheets' tab close).
 * Death hooks are attached only while a sender has in-flight streams and are
 * detached again once its last stream settles.
 */
export class RendererAiStreams {
  private readonly bySender = new Map<number, { registry: AiStreamRegistry; detach: () => void }>()
  private readonly byRequest = new Map<string, AbortController>()

  /** Registers the abort controller of one starting stream for its sender. */
  start(sender: AiStreamSender, requestId: string): AbortController {
    let entry = this.bySender.get(sender.id)
    if (!entry) {
      const registry = new AiStreamRegistry()
      const detach = (): void => {
        sender.removeListener('destroyed', onGone)
        sender.removeListener('render-process-gone', onGone)
      }
      const onGone = (): void => {
        detach()
        this.bySender.delete(sender.id)
        registry.abortAll()
      }
      entry = { registry, detach }
      this.bySender.set(sender.id, entry)
      sender.once('destroyed', onGone)
      sender.once('render-process-gone', onGone)
    }
    const controller = entry.registry.start(requestId)
    this.byRequest.set(requestId, controller)
    return controller
  }

  /**
   * Drops a settled stream (done, error, or abort); detaches the sender's
   * death hooks once it has no in-flight streams left.
   */
  finish(sender: AiStreamSender, requestId: string): void {
    this.byRequest.delete(requestId)
    const entry = this.bySender.get(sender.id)
    if (!entry) return
    entry.registry.finish(requestId)
    if (entry.registry.size === 0) {
      this.bySender.delete(sender.id)
      entry.detach()
    }
  }

  /** Aborts one in-flight stream (renderer cancel request); unknown ids are a no-op. */
  cancel(requestId: string): void {
    this.byRequest.get(requestId)?.abort()
  }
}

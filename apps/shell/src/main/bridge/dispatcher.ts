// FIFO request dispatcher for the Airy live bridge. Requests are handled one
// at a time, strictly in arrival order — one local client issuing sequential
// calls needs no id correlation on the wire (PLAN.md §2). Pure Node module,
// unit-testable without Electron (tests/bridge-server.test.ts).
import {
  BRIDGE_PROTOCOL_VERSION,
  BridgeMethodError,
  bridgeError,
  type BridgeRequest,
  type BridgeResponse,
} from './protocol'

export type BridgeMethodHandler = (
  params: Record<string, unknown>,
  /** aborts when the per-call timeout fires, so round-trips can drop waiters */
  signal: AbortSignal,
  /** per-connection context the server stamps onto every dispatched call */
  context: BridgeCallContext,
) => unknown | Promise<unknown>

/** identity of the calling bridge connection, stable for its lifetime */
export interface BridgeCallContext {
  readonly clientId: string
}

/** fallback for direct dispatcher callers (tests) with no connection identity */
export const LOCAL_BRIDGE_CALLER: BridgeCallContext = { clientId: 'local' }

export interface BridgeDispatcher {
  call(request: BridgeRequest, context?: BridgeCallContext): Promise<BridgeResponse>
}

export const DEFAULT_BRIDGE_TIMEOUT_MS = 30_000

function withTimeout(
  run: (signal: AbortSignal) => Promise<unknown>,
  timeoutMs: number,
  onTimeout: () => void,
): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const controller = new AbortController()
    const timer = setTimeout(() => {
      // abort first so the handler can drop its waiter while the caller waits
      controller.abort()
      onTimeout()
      reject(new BridgeMethodError('timeout', `timed out after ${timeoutMs}ms`))
    }, timeoutMs)
    run(controller.signal).then(
      (value) => {
        clearTimeout(timer)
        resolve(value)
      },
      (err) => {
        clearTimeout(timer)
        reject(err)
      },
    )
  })
}

export function createBridgeDispatcher(
  methods: Record<string, BridgeMethodHandler>,
  options: { timeoutMs?: number } = {},
): BridgeDispatcher {
  const timeoutMs = options.timeoutMs ?? DEFAULT_BRIDGE_TIMEOUT_MS
  return {
    async call(request: BridgeRequest, context?: BridgeCallContext): Promise<BridgeResponse> {
      // parseRequestLine already screens this; re-check so a dispatcher wired
      // to another transport stays version-safe
      if (request.protocol_version !== BRIDGE_PROTOCOL_VERSION) {
        return bridgeError(
          'unsupported_version',
          `protocol_version ${String(request.protocol_version)} is not supported (server speaks ${BRIDGE_PROTOCOL_VERSION})`,
        )
      }
      const handler = methods[request.method]
      if (!handler) {
        return bridgeError(
          'unknown_method',
          `unknown method "${request.method}" (supported: ${Object.keys(methods)
            .sort()
            .join(', ')})`,
        )
      }
      try {
        const result = await withTimeout(
          (signal) =>
            Promise.resolve(handler(request.params ?? {}, signal, context ?? LOCAL_BRIDGE_CALLER)),
          timeoutMs,
          () => {},
        )
        return { ok: true, result }
      } catch (err) {
        if (err instanceof BridgeMethodError) {
          return bridgeError(err.code, err.message)
        }
        const message = err instanceof Error ? err.message : String(err)
        return bridgeError('internal', message)
      }
    },
  }
}

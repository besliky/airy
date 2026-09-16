// Client for the Airy live bridge: connects to the running app's UDS/named-
// pipe server, performs the token handshake, and executes bridge calls one at
// a time in strict FIFO order (the protocol has no request ids — responses are
// matched by arrival order, which equals request order because the server is
// FIFO and this client never has more than one request in flight). No auto-
// reconnect: a failed call drops the connection and the next call connects
// fresh, rereading the info file so a restarted app's new token is picked up.
import { connect as netConnect, type Socket } from 'node:net'

import {
  candidateBridgeInfoPaths,
  discoverBridgeInfos,
  type DiscoveryOptions,
  type DiscoveredBridge,
} from './discovery.js'
import {
  BRIDGE_PROTOCOL_VERSION,
  NdjsonFramer,
  parseResponseLine,
  type BridgeErrorCode,
  type BridgeResponse,
} from './protocol.js'

/** default per-call budget, mirroring the server's dispatcher timeout */
export const DEFAULT_BRIDGE_CALL_TIMEOUT_MS = 30_000
const CONNECT_TIMEOUT_MS = 5_000

export type LiveBridgeErrorCode = 'bridge_not_running' | 'bridge_unauthorized' | 'bridge_error'

/**
 * Client-side failure with a stable code:
 * - bridge_not_running  — no info file, unreachable socket, connection lost
 * - bridge_unauthorized — the handshake token was rejected
 * - bridge_error        — the server answered {ok:false,error{code,message}}
 *                         (bridgeCode carries the protocol-level code) or the
 *                         response stream was malformed/desynced
 */
export class BridgeClientError extends Error {
  readonly code: LiveBridgeErrorCode
  /** protocol-level error code when the failure originated server-side */
  readonly bridgeCode?: BridgeErrorCode

  constructor(code: LiveBridgeErrorCode, message: string, bridgeCode?: BridgeErrorCode) {
    super(message)
    this.name = 'BridgeClientError'
    this.code = code
    this.bridgeCode = bridgeCode
  }
}

function serverError(response: Extract<BridgeResponse, { ok: false }>): BridgeClientError {
  const error = response.error
  return new BridgeClientError('bridge_error', `${error.code}: ${error.message}`, error.code)
}

/**
 * One live connection: frames NDJSON input and hands out response lines to a
 * single waiting reader (calls are serialized, so one slot suffices). Any
 * socket error, close, or read timeout kills the connection — a desynced FIFO
 * stream must never be reused.
 */
class BridgeConnection {
  private framer = new NdjsonFramer()
  private buffered: string[] = []
  private reader: { resolve: (line: string) => void; reject: (err: Error) => void } | null = null
  private failure: Error | null = null

  constructor(
    readonly socket: Socket,
    private readonly onDead: (connection: BridgeConnection) => void,
  ) {
    socket.on('data', (chunk: Buffer) => this.onData(chunk))
    socket.on('error', (err: Error) =>
      this.die(
        new BridgeClientError('bridge_not_running', `bridge connection error: ${err.message}`),
      ),
    )
    socket.on('close', () =>
      this.die(new BridgeClientError('bridge_not_running', 'the bridge connection was lost')),
    )
  }

  get dead(): boolean {
    return this.failure !== null
  }

  private onData(chunk: Buffer): void {
    const { lines, overflow } = this.framer.push(chunk)
    if (overflow) {
      this.die(
        new BridgeClientError(
          'bridge_error',
          `the bridge response exceeded the ${this.framer.maxLineBytes}-byte line limit`,
        ),
      )
      return
    }
    for (const line of lines) {
      const reader = this.reader
      this.reader = null
      if (reader) reader.resolve(line)
      else this.buffered.push(line)
    }
  }

  private die(err: BridgeClientError): void {
    if (this.failure) return
    this.failure = err
    this.socket.destroy()
    const reader = this.reader
    this.reader = null
    reader?.reject(err)
    this.onDead(this)
  }

  send(request: object): void {
    this.socket.write(`${JSON.stringify(request)}\n`)
  }

  /** drop the socket without touching pending waiters (they reject via 'close') */
  close(): void {
    this.socket.destroy()
  }

  /** await the next response line; rejects (and kills the connection) on timeout */
  readLine(timeoutMs: number, context: string): Promise<string> {
    const buffered = this.buffered.shift()
    if (buffered !== undefined) return Promise.resolve(buffered)
    if (this.failure) return Promise.reject(this.failure)
    return new Promise<string>((resolve, reject) => {
      const timer = setTimeout(() => {
        // no ids on the wire: a late reply would answer the NEXT call — drop
        // the connection so the next call reconnects cleanly
        this.die(
          new BridgeClientError(
            'bridge_error',
            `${context} timed out after ${timeoutMs}ms`,
            'timeout',
          ),
        )
      }, timeoutMs)
      this.reader = {
        resolve: (line) => {
          clearTimeout(timer)
          resolve(line)
        },
        reject: (err) => {
          clearTimeout(timer)
          reject(err)
        },
      }
    })
  }
}

export interface LiveBridge {
  /**
   * One FIFO bridge call: connect (lazily) + handshake if needed, send the
   * request, await the response. Rejects with BridgeClientError on failure.
   */
  call(method: string, params?: Record<string, unknown>): Promise<unknown>
  /** drop the connection (the next call reconnects) */
  close(): void
}

export function createLiveBridge(
  options: { timeoutMs?: number } & DiscoveryOptions = {},
): LiveBridge {
  const timeoutMs = options.timeoutMs ?? DEFAULT_BRIDGE_CALL_TIMEOUT_MS
  const discoveryOptions: DiscoveryOptions = {
    env: options.env,
    platform: options.platform,
    homeDir: options.homeDir,
    isProcessAlive: options.isProcessAlive,
  }
  let connection: BridgeConnection | null = null
  let chain: Promise<unknown> = Promise.resolve()

  function notRunning(): BridgeClientError {
    const searched = candidateBridgeInfoPaths(discoveryOptions)
      .map((p) => `\n  - ${p}`)
      .join('')
    return new BridgeClientError(
      'bridge_not_running',
      `no bridge info file found (searched:${searched})`,
    )
  }

  /** parse a response line; malformed output kills the connection (desync) */
  function parseOrKill(conn: BridgeConnection, line: string, context: string): BridgeResponse {
    const parsed = parseResponseLine(line)
    if (!parsed.ok) {
      conn.close()
      throw new BridgeClientError(
        'bridge_error',
        `${context}: malformed response (${parsed.message})`,
      )
    }
    return parsed.value
  }

  /** connect + handshake one candidate */
  async function connectCandidate(found: DiscoveredBridge): Promise<BridgeConnection> {
    const socket = netConnect(found.info.socketPath)
    await new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => {
        socket.destroy()
        reject(
          new BridgeClientError(
            'bridge_not_running',
            `timed out connecting to the bridge socket ${found.info.socketPath}`,
          ),
        )
      }, CONNECT_TIMEOUT_MS)
      socket.once('connect', () => {
        clearTimeout(timer)
        resolve()
      })
      socket.once('error', (err: Error) => {
        clearTimeout(timer)
        // ENOENT: stale socket file; ECONNREFUSED: dead socket — either way
        // the app side is not serving
        reject(
          new BridgeClientError(
            'bridge_not_running',
            `cannot connect to the bridge socket ${found.info.socketPath}: ${err.message}`,
          ),
        )
      })
    })

    const conn = new BridgeConnection(socket, (dead) => {
      if (connection === dead) connection = null
    })
    // handshake: hello carrying the info file's token must be the first
    // request; the client just presents the token (timing-safe comparison is
    // the server's concern)
    conn.send({
      protocol_version: BRIDGE_PROTOCOL_VERSION,
      method: 'hello',
      params: { token: found.info.token },
    })
    const line = await conn.readLine(timeoutMs, 'bridge handshake')
    const response = parseOrKill(conn, line, 'bridge handshake')
    if (!response.ok) {
      conn.close()
      throw response.error.code === 'unauthorized'
        ? new BridgeClientError('bridge_unauthorized', response.error.message)
        : serverError(response)
    }
    connection = conn
    return conn
  }

  async function connectAndHandshake(): Promise<BridgeConnection> {
    const candidates = await discoverBridgeInfos(discoveryOptions)
    // a candidate whose socket is unreachable (stale file, crashed app) must
    // not strand the later ones: connect failures fall through to the next
    // candidate; only the last error surfaces. Handshake-level failures
    // (unauthorized, malformed stream) belong to a live server and are thrown.
    let lastConnectError: BridgeClientError | null = null
    for (const found of candidates) {
      try {
        return await connectCandidate(found)
      } catch (err) {
        if (err instanceof BridgeClientError && err.code === 'bridge_not_running') {
          lastConnectError = err
          continue
        }
        throw err
      }
    }
    throw lastConnectError ?? notRunning()
  }

  async function performCall(method: string, params?: Record<string, unknown>): Promise<unknown> {
    const conn = connection && !connection.dead ? connection : await connectAndHandshake()
    conn.send({
      protocol_version: BRIDGE_PROTOCOL_VERSION,
      method,
      ...(params !== undefined ? { params } : {}),
    })
    const line = await conn.readLine(timeoutMs, `bridge call "${method}"`)
    const response = parseOrKill(conn, line, `bridge call "${method}"`)
    if (!response.ok) throw serverError(response)
    return response.result
  }

  return {
    call(method: string, params?: Record<string, unknown>): Promise<unknown> {
      const run = () => performCall(method, params)
      // strictly one in-flight call; a failure must not stall later callers
      const next = chain.then(run, run)
      chain = next.catch(() => {})
      return next
    },
    close(): void {
      connection?.socket.destroy()
      connection = null
    },
  }
}

let shared: LiveBridge | null = null

/** the process-wide client used by the live_* tools (keeps calls FIFO) */
export function sharedLiveBridge(): LiveBridge {
  shared ??= createLiveBridge()
  return shared
}

/** tests: drop the shared connection so the next call reconnects from scratch */
export function resetSharedLiveBridge(): void {
  shared?.close()
  shared = null
}

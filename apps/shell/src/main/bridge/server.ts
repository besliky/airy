// Transport for the Airy live bridge: a Unix domain socket (linux/mac) or a
// Windows named pipe inside userData, plus a 0600 info file carrying the
// socket path and the per-session token. Pure Node (node:net + node:fs) —
// Electron only supplies the userData directory, so this module is testable
// in plain Node against a temp directory (tests/bridge-server.test.ts).
import { randomBytes } from 'node:crypto'
import { chmod, rm, writeFile } from 'node:fs/promises'
import { createServer as createNetServer, type Socket } from 'node:net'
import { join } from 'node:path'

import { createBridgeDispatcher, type BridgeMethodHandler } from './dispatcher'
import {
  BRIDGE_PROTOCOL_VERSION,
  bridgeError,
  createHandshakeGate,
  encodeResponse,
  NdjsonFramer,
  parseRequestLine,
  type BridgeResponse,
} from './protocol'

export const BRIDGE_SOCKET_NAME = 'airy-bridge.sock'
export const BRIDGE_PIPE_PATH = '\\\\.\\pipe\\airy-bridge'
export const BRIDGE_INFO_NAME = 'airy-bridge.json'

/** endpoint description published next to the socket for local clients */
export interface BridgeEndpointInfo {
  socketPath: string
  token: string
  pid: number
  protocolVersion: number
}

export function bridgeSocketPath(
  userDataDir: string,
  platform: NodeJS.Platform = process.platform,
) {
  return platform === 'win32' ? BRIDGE_PIPE_PATH : join(userDataDir, BRIDGE_SOCKET_NAME)
}

export function bridgeInfoPath(userDataDir: string): string {
  return join(userDataDir, BRIDGE_INFO_NAME)
}

export function generateBridgeToken(): string {
  return randomBytes(32).toString('hex')
}

/// The socket surface the backpressure writer needs (satisfied by net.Socket,
/// faked in tests).
export interface BackpressureSocket {
  write(payload: string): boolean
  pause(): void
  resume(): void
  on(event: 'drain', listener: () => void): void
}

/**
 * Write-side backpressure for one connection. A peer that reads slowly makes
 * socket.write() return false once the kernel buffer fills; writing on
 * regardless would buffer responses in memory without bound. Pause the
 * socket — request ingress stops with it — until the buffer drains.
 */
export function createBackpressureWriter(socket: BackpressureSocket): {
  write(payload: string): void
  readonly paused: boolean
} {
  let paused = false
  socket.on('drain', () => {
    if (!paused) return
    paused = false
    socket.resume()
  })
  return {
    get paused() {
      return paused
    },
    write(payload) {
      if (socket.write(payload)) return
      paused = true
      socket.pause()
    },
  }
}

/**
 * Write the info file and tighten it to 0600 — Node creates files 0775 &
  umask, and the token grants full document-edit access, so owner-only is the
 * baseline (PLAN.md Phase 2). Windows ignores POSIX mode bits; the pipe and
 * file already live under the user profile.
 */
export async function writeBridgeInfoFile(path: string, info: BridgeEndpointInfo): Promise<void> {
  await writeFile(path, `${JSON.stringify(info, null, 2)}\n`, { mode: 0o600 })
  if (process.platform !== 'win32') await chmod(path, 0o600)
}

export interface BridgeServerHandle {
  info: BridgeEndpointInfo
  /** close the listener, tear down live connections, remove socket + info file */
  stop(): Promise<void>
}

export async function startBridgeServer(options: {
  userDataDir: string
  methods: Record<string, BridgeMethodHandler>
  timeoutMs?: number
  /** requests a connection may pipeline before it is closed as abusive */
  maxQueuedRequests?: number
  log?: (message: string) => void
}): Promise<BridgeServerHandle> {
  const { userDataDir, methods, timeoutMs, log = () => {} } = options
  const maxQueuedRequests = options.maxQueuedRequests ?? 256
  const socketPath = bridgeSocketPath(userDataDir)
  const infoPath = bridgeInfoPath(userDataDir)
  const token = generateBridgeToken()
  // hello is always available: the handshake gate authorizes the connection,
  // this handler answers it (a caller-supplied hello cannot override it)
  const methodMap: Record<string, BridgeMethodHandler> = {
    hello: () => ({
      server: 'airy-bridge',
      protocolVersion: BRIDGE_PROTOCOL_VERSION,
      pid: process.pid,
    }),
  }
  for (const [name, handler] of Object.entries(methods)) {
    if (name !== 'hello') methodMap[name] = handler
  }
  const dispatcher = createBridgeDispatcher(methodMap, { timeoutMs })
  const clients = new Set<Socket>()
  // per-server connection counter: each socket gets a stable identity for the
  // lifetime of the connection (bridge turn ownership in the renderer)
  let nextConnectionId = 0

  const server = createNetServer((socket) => {
    clients.add(socket)
    const gate = createHandshakeGate(token)
    const framer = new NdjsonFramer()
    const clientId = `conn-${(nextConnectionId += 1)}`
    let closed = false
    // responses are chained per connection: request N+1 only starts after N's
    // response was written — the FIFO contract, even when handlers are async
    let chain: Promise<void> = Promise.resolve()
    // requests accepted but not yet answered; a peer pipelining faster than
    // it reads would grow this (and the socket buffer) without bound
    let queued = 0
    const wire = createBackpressureWriter(socket)
    const write = (response: BridgeResponse) => {
      if (closed || socket.destroyed) return
      const payload = encodeResponse(response)
      // BUG-904: an oversized result would trip the client framer's line cap
      // and kill the connection with no protocol-level error (get_context of
      // a huge document). Replace it with a typed invalid_request — the FIFO
      // stream stays in sync (the call itself already ran) and the client
      // learns to narrow the request instead of reconnecting blind.
      if (Buffer.byteLength(payload) > framer.maxLineBytes) {
        wire.write(
          `${encodeResponse(
            bridgeError(
              'invalid_request',
              `response exceeds the ${framer.maxLineBytes}-byte line limit (result too large for the bridge)`,
            ),
          )}\n`,
        )
        return
      }
      wire.write(`${payload}\n`)
    }
    const close = () => {
      if (closed) return
      closed = true
      // end (not destroy): the pending error response must still flush; the
      // destroy in the end callback tears the read side down afterwards so a
      // peer that keeps writing cannot hold the half-open socket (and its
      // framer buffer) alive indefinitely
      socket.end(() => socket.destroy())
      clients.delete(socket)
    }
    socket.on('data', (chunk: Buffer) => {
      // post-close data must not dispatch: close() only half-closes the
      // socket, so data events keep arriving while the error flushes — a
      // request executed there would apply its mutation with the response
      // silently dropped, and a retrying client would apply it twice
      // (BUG-901: line overflow, queue cap, handshake reject all close)
      if (closed) return
      const { lines, overflow } = framer.push(chunk)
      if (overflow) {
        write(
          bridgeError(
            'invalid_request',
            `request exceeds the ${framer.maxLineBytes}-byte line limit`,
          ),
        )
        close()
        return
      }
      for (const line of lines) {
        const parsed = parseRequestLine(line)
        if (!parsed.ok) {
          write(bridgeError(parsed.error.code, parsed.error.message))
          continue
        }
        const rejected = gate.admit(parsed.value)
        if (rejected) {
          write(bridgeError(rejected.code, rejected.message))
          // no valid handshake: nothing beyond this point is trusted
          close()
          return
        }
        const request = parsed.value
        queued += 1
        // The queue is unbounded work in memory while the earlier responses
        // wait on a stalled or slow-reading peer; refuse the flood outright.
        // The rejection jumps the FIFO chain on purpose — the connection is
        // being closed, and the reason must reach the peer before it does.
        if (queued > maxQueuedRequests) {
          write(
            bridgeError(
              'invalid_request',
              `too many pipelined requests (over ${maxQueuedRequests}) — closing`,
            ),
          )
          close()
          return
        }
        chain = chain
          .then(() => dispatcher.call(request, { clientId }))
          .then(
            (response) => {
              write(response)
              queued -= 1
            },
            (err: unknown) => {
              // dispatcher.call never rejects, but a write failure must not break the chain
              write(
                bridgeError(
                  'internal',
                  err instanceof Error ? err.message : 'unexpected bridge failure',
                ),
              )
              queued -= 1
            },
          )
      }
    })
    socket.on('error', (err) => log(`bridge connection error: ${String(err)}`))
    socket.on('close', () => {
      closed = true
      clients.delete(socket)
    })
  })

  // a previous instance may have died without unlinking its socket
  if (process.platform !== 'win32') await rm(socketPath, { force: true })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(socketPath, () => resolve())
  })
  // Node creates listening sockets 0775 & umask — restrict to the owner, like
  // the info file (chmod explicitly: PLAN.md, SB8 finding)
  if (process.platform !== 'win32') await chmod(socketPath, 0o600)

  const info: BridgeEndpointInfo = {
    socketPath,
    token,
    pid: process.pid,
    protocolVersion: BRIDGE_PROTOCOL_VERSION,
  }
  await writeBridgeInfoFile(infoPath, info)
  log(`bridge listening on ${socketPath}`)

  return {
    info,
    async stop() {
      await new Promise<void>((resolve) => {
        server.close(() => resolve())
        for (const client of clients) client.destroy()
        clients.clear()
      })
      await rm(infoPath, { force: true })
      if (process.platform !== 'win32') await rm(socketPath, { force: true })
    },
  }
}

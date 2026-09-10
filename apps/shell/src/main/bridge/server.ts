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
  log?: (message: string) => void
}): Promise<BridgeServerHandle> {
  const { userDataDir, methods, timeoutMs, log = () => {} } = options
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

  const server = createNetServer((socket) => {
    clients.add(socket)
    const gate = createHandshakeGate(token)
    const framer = new NdjsonFramer()
    let closed = false
    // responses are chained per connection: request N+1 only starts after N's
    // response was written — the FIFO contract, even when handlers are async
    let chain: Promise<void> = Promise.resolve()
    const write = (response: BridgeResponse) => {
      if (closed || socket.destroyed) return
      socket.write(`${encodeResponse(response)}\n`)
    }
    const close = () => {
      if (closed) return
      closed = true
      // end (not destroy): the pending error response must still flush
      socket.end()
      clients.delete(socket)
    }
    socket.on('data', (chunk: Buffer) => {
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
        chain = chain
          .then(() => dispatcher.call(request))
          .then(write, (err: unknown) => {
            // dispatcher.call never rejects, but a write failure must not break the chain
            write(
              bridgeError(
                'internal',
                err instanceof Error ? err.message : 'unexpected bridge failure',
              ),
            )
          })
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

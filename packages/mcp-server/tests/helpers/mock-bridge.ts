// A real bridge server for tests: UDS in a temp directory + the 0600 info
// file, exactly per the app-side contract (apps/shell/src/main/bridge/
// server.ts) — hello-gated handshake, FIFO NDJSON dispatch, per-call timeout.
// The live_* tool tests point AIRY_BRIDGE_FILE at its info file, so the MCP
// client runs against a genuine socket + token flow, not a stub.
import { randomBytes } from 'node:crypto'
import { chmod, rm, writeFile } from 'node:fs/promises'
import { createServer as createNetServer, type Socket } from 'node:net'
import { join } from 'node:path'

import {
  BRIDGE_PROTOCOL_VERSION,
  isBridgeErrorCode,
  NdjsonFramer,
  parseRequestLine,
  type BridgeResponse,
} from '../../src/live/protocol.js'
import type { BridgeEndpointInfo } from '../../src/live/discovery.js'

/** handler result; throwing {code,message} shapes a protocol error response */
export type MockBridgeHandler = (params: Record<string, unknown>) => unknown | Promise<unknown>

export class MockBridgeMethodError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message)
    this.name = 'MockBridgeMethodError'
  }
}

async function writeInfo(path: string, info: BridgeEndpointInfo): Promise<void> {
  await writeFile(path, `${JSON.stringify(info, null, 2)}\n`, { mode: 0o600 })
  if (process.platform !== 'win32') await chmod(path, 0o600)
}

export interface MockBridgeHandle {
  info: BridgeEndpointInfo
  infoPath: string
  socketPath: string
  /** admitted (post-handshake) business requests, in arrival order */
  requests: Array<{ method: string; params: Record<string, unknown> }>
  /** rewrite the info file the client reads (e.g. a wrong token) */
  rewriteInfo(patch: Partial<BridgeEndpointInfo>): Promise<void>
  stop(): Promise<void>
}

export async function startMockBridge(options: {
  dir: string
  methods: Record<string, MockBridgeHandler>
  timeoutMs?: number
}): Promise<MockBridgeHandle> {
  const { dir, methods, timeoutMs = 1_000 } = options
  const socketPath = join(dir, 'airy-bridge.sock')
  const infoPath = join(dir, 'airy-bridge.json')
  const token = randomBytes(32).toString('hex')
  const requests: Array<{ method: string; params: Record<string, unknown> }> = []
  const clients = new Set<Socket>()

  // dispatcher: run the handler under a timeout, map throws to error envelopes
  const dispatch = async (
    method: string,
    params: Record<string, unknown>,
  ): Promise<BridgeResponse> => {
    const handler = method === 'hello' ? null : methods[method]
    if (!handler) {
      return {
        ok: false,
        error: {
          code: 'unknown_method',
          message: `unknown method "${method}" (supported: ${['hello', ...Object.keys(methods)].sort().join(', ')})`,
        },
      }
    }
    try {
      const result = await new Promise<unknown>((resolve, reject) => {
        const timer = setTimeout(
          () => reject(new MockBridgeMethodError('timeout', `timed out after ${timeoutMs}ms`)),
          timeoutMs,
        )
        Promise.resolve(handler(params)).then(
          (value) => {
            clearTimeout(timer)
            resolve(value)
          },
          (err: unknown) => {
            clearTimeout(timer)
            reject(err)
          },
        )
      })
      return { ok: true, result }
    } catch (err) {
      if (err instanceof MockBridgeMethodError && isBridgeErrorCode(err.code)) {
        return { ok: false, error: { code: err.code, message: err.message } }
      }
      return {
        ok: false,
        error: { code: 'internal', message: err instanceof Error ? err.message : String(err) },
      }
    }
  }

  const server = createNetServer((socket) => {
    clients.add(socket)
    const framer = new NdjsonFramer()
    let authorized = false
    let closed = false
    // responses chained per connection — the FIFO contract
    let chain: Promise<void> = Promise.resolve()
    const write = (response: BridgeResponse) => {
      if (closed || socket.destroyed) return
      socket.write(`${JSON.stringify(response)}\n`)
    }
    const close = () => {
      if (closed) return
      closed = true
      socket.end()
      clients.delete(socket)
    }
    socket.on('data', (chunk: Buffer) => {
      const { lines } = framer.push(chunk)
      for (const line of lines) {
        const parsed = parseRequestLine(line)
        if (!parsed.ok) {
          write({ ok: false, error: parsed.error })
          continue
        }
        const request = parsed.value
        // hello gate: the first request must carry the token
        if (!authorized) {
          if (request.method !== 'hello') {
            write({
              ok: false,
              error: {
                code: 'unauthorized',
                message: 'the first request must be hello with the bridge token',
              },
            })
            close()
            return
          }
          const presented = (request.params as { token?: unknown } | undefined)?.token
          if (presented !== token) {
            write({ ok: false, error: { code: 'unauthorized', message: 'invalid bridge token' } })
            close()
            return
          }
          authorized = true
          write({
            ok: true,
            result: {
              server: 'airy-bridge',
              protocolVersion: BRIDGE_PROTOCOL_VERSION,
              pid: info.pid,
            },
          })
          continue
        }
        if (request.method === 'hello') {
          write({
            ok: false,
            error: { code: 'invalid_request', message: 'hello is only valid as the first message' },
          })
          continue
        }
        requests.push({ method: request.method, params: request.params ?? {} })
        chain = chain
          .then(() => dispatch(request.method, request.params ?? {}))
          .then(write, (err: unknown) => {
            write({
              ok: false,
              error: {
                code: 'internal',
                message: err instanceof Error ? err.message : String(err),
              },
            })
          })
      }
    })
    socket.on('error', () => {})
    socket.on('close', () => {
      closed = true
      clients.delete(socket)
    })
  })

  const info: BridgeEndpointInfo = {
    socketPath,
    token,
    pid: process.pid,
    protocolVersion: BRIDGE_PROTOCOL_VERSION,
  }

  await rm(socketPath, { force: true })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(socketPath, () => resolve())
  })
  if (process.platform !== 'win32') await chmod(socketPath, 0o600)
  await writeInfo(infoPath, info)

  return {
    info,
    infoPath,
    socketPath,
    requests,
    rewriteInfo(patch: Partial<BridgeEndpointInfo>) {
      return writeInfo(infoPath, { ...info, ...patch })
    },
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

import { existsSync, statSync } from 'node:fs'
import { mkdtempSync, rmSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { connect } from 'node:net'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import {
  bridgeInfoPath,
  bridgeSocketPath,
  BRIDGE_INFO_NAME,
  BRIDGE_SOCKET_NAME,
  createBackpressureWriter,
  startBridgeServer,
  writeBridgeInfoFile,
  type BackpressureSocket,
  type BridgeEndpointInfo,
  type BridgeServerHandle,
} from '../src/main/bridge/server'

/**
 * Bridge transport (src/main/bridge/server.ts) over a real Unix domain socket
 * in a temp dir: token-file permissions, the hello handshake, FIFO response
 * ordering, and per-call timeouts — all in plain Node, no Electron.
 */

let dir: string
let server: BridgeServerHandle | null = null

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'airy-bridge-test-'))
})

afterEach(async () => {
  await server?.stop()
  server = null
  rmSync(dir, { recursive: true, force: true })
})

interface ClientLine {
  resolve(line: string): void
}

/** minimal NDJSON client: writes request lines, collects response lines */
function connectClient(socketPath: string) {
  const socket = connect(socketPath)
  const pending: ClientLine[] = []
  const lines: string[] = []
  let buffer = ''
  const closed = new Promise<void>((resolve) => socket.on('close', () => resolve()))
  socket.on('data', (chunk: Buffer) => {
    buffer += chunk.toString('utf8')
    let newline = buffer.indexOf('\n')
    while (newline !== -1) {
      const line = buffer.slice(0, newline).replace(/\r$/, '')
      buffer = buffer.slice(newline + 1)
      // a waiting reader consumes the line directly; otherwise buffer it
      const waiter = pending.shift()
      if (waiter) waiter.resolve(line)
      else lines.push(line)
      newline = buffer.indexOf('\n')
    }
  })
  const ready = new Promise<void>((resolve, reject) => {
    socket.once('connect', () => resolve())
    socket.once('error', reject)
  })
  return {
    ready,
    closed,
    send(payload: unknown): void {
      socket.write(`${JSON.stringify(payload)}\n`)
    },
    /** raw write for flood-shaped payloads the JSON helper would chunk one-by-one */
    writeRaw(text: string): void {
      socket.write(text)
    },
    async next(): Promise<string> {
      const line = lines.shift()
      if (line !== undefined) return line
      return new Promise((resolve) => pending.push({ resolve }))
    },
    end(): void {
      socket.end()
    },
  }
}

function handshake(client: ReturnType<typeof connectClient>, token: string): Promise<string> {
  client.send({ protocol_version: 1, method: 'hello', params: { token } })
  return client.next()
}

describe('path + token file generation', () => {
  it('places socket and info file inside userData', () => {
    expect(bridgeSocketPath('/ud')).toBe(join('/ud', BRIDGE_SOCKET_NAME))
    expect(bridgeSocketPath('/ud', 'win32')).toBe('\\\\.\\pipe\\airy-bridge')
    expect(bridgeInfoPath('/ud')).toBe(join('/ud', BRIDGE_INFO_NAME))
  })

  it('writes the info file with 0600 permissions', async () => {
    if (process.platform === 'win32') return
    const info: BridgeEndpointInfo = {
      socketPath: '/tmp/s.sock',
      token: 't'.repeat(64),
      pid: 4242,
      protocolVersion: 1,
    }
    const path = join(dir, 'info.json')
    await writeBridgeInfoFile(path, info)
    expect(JSON.parse(await readFile(path, 'utf8'))).toEqual(info)
    // Node's default umask yields 0775-ish files; the token file must be owner-only
    expect(statSync(path).mode & 0o777).toBe(0o600)
  })
})

describe('bridge server over a live socket', () => {
  it('publishes socketPath/token/pid in the 0600 info file', async () => {
    if (process.platform === 'win32') return
    server = await startBridgeServer({ userDataDir: dir, methods: { ping: () => 1 } })
    const infoPath = bridgeInfoPath(dir)
    expect(existsSync(infoPath)).toBe(true)
    const info = JSON.parse(await readFile(infoPath, 'utf8')) as BridgeEndpointInfo
    expect(info.socketPath).toBe(server.info.socketPath)
    expect(info.token).toBe(server.info.token)
    expect(info.pid).toBe(process.pid)
    expect(info.protocolVersion).toBe(1)
    expect(statSync(infoPath).mode & 0o777).toBe(0o600)
    expect(statSync(server.info.socketPath).mode & 0o777).toBe(0o600)
  })

  it('answers hello + ping', async () => {
    server = await startBridgeServer({
      userDataDir: dir,
      methods: { ping: () => ({ pong: true }) },
    })
    const client = connectClient(server.info.socketPath)
    await client.ready
    const hello = JSON.parse(await handshake(client, server.info.token))
    expect(hello).toEqual({
      ok: true,
      result: { server: 'airy-bridge', protocolVersion: 1, pid: process.pid },
    })
    client.send({ protocol_version: 1, method: 'ping' })
    expect(JSON.parse(await client.next())).toEqual({ ok: true, result: { pong: true } })
    client.end()
  })

  it('closes the connection when the first message is not hello', async () => {
    server = await startBridgeServer({ userDataDir: dir, methods: { ping: () => 1 } })
    const client = connectClient(server.info.socketPath)
    await client.ready
    client.send({ protocol_version: 1, method: 'ping' })
    expect(JSON.parse(await client.next())).toEqual({
      ok: false,
      error: {
        code: 'unauthorized',
        message: 'the first request must be hello with the bridge token',
      },
    })
    await client.closed
  })

  it('rejects a wrong token and closes', async () => {
    server = await startBridgeServer({ userDataDir: dir, methods: {} })
    const client = connectClient(server.info.socketPath)
    await client.ready
    expect(JSON.parse(await handshake(client, 'wrong-token'))).toEqual({
      ok: false,
      error: { code: 'unauthorized', message: 'invalid bridge token' },
    })
    await client.closed
  })

  it('answers unknown_method and unsupported_version without closing', async () => {
    server = await startBridgeServer({ userDataDir: dir, methods: { ping: () => 1 } })
    const client = connectClient(server.info.socketPath)
    await client.ready
    await handshake(client, server.info.token)
    client.send({ protocol_version: 1, method: 'frobnicate' })
    expect(JSON.parse(await client.next())).toEqual({
      ok: false,
      error: {
        code: 'unknown_method',
        message: 'unknown method "frobnicate" (supported: hello, ping)',
      },
    })
    client.send({ protocol_version: 2, method: 'ping' })
    expect(JSON.parse(await client.next())).toEqual({
      ok: false,
      error: {
        code: 'unsupported_version',
        message: 'protocol_version 2 is not supported (server speaks 1)',
      },
    })
    // still alive: ping works after the errors
    client.send({ protocol_version: 1, method: 'ping' })
    expect(JSON.parse(await client.next())).toEqual({ ok: true, result: 1 })
    client.end()
  })

  it('keeps responses strictly FIFO when handlers resolve out of order', async () => {
    server = await startBridgeServer({
      userDataDir: dir,
      methods: {
        slow: () => new Promise((resolve) => setTimeout(() => resolve('slow-result'), 80)),
        fast: () => 'fast-result',
      },
    })
    const client = connectClient(server.info.socketPath)
    await client.ready
    await handshake(client, server.info.token)
    client.send({ protocol_version: 1, method: 'slow' })
    client.send({ protocol_version: 1, method: 'fast' })
    expect(JSON.parse(await client.next())).toEqual({ ok: true, result: 'slow-result' })
    expect(JSON.parse(await client.next())).toEqual({ ok: true, result: 'fast-result' })
    client.end()
  })

  it('times a stalled call out and keeps serving after it', async () => {
    server = await startBridgeServer({
      userDataDir: dir,
      timeoutMs: 60,
      methods: {
        stall: (_params, signal) =>
          new Promise((_resolve, reject) =>
            signal.addEventListener('abort', () => reject(new Error('aborted'))),
          ),
        ping: () => 'still-alive',
      },
    })
    const client = connectClient(server.info.socketPath)
    await client.ready
    await handshake(client, server.info.token)
    client.send({ protocol_version: 1, method: 'stall' })
    expect(JSON.parse(await client.next())).toEqual({
      ok: false,
      error: { code: 'timeout', message: 'timed out after 60ms' },
    })
    client.send({ protocol_version: 1, method: 'ping' })
    expect(JSON.parse(await client.next())).toEqual({ ok: true, result: 'still-alive' })
    client.end()
  })

  it('serves two independent clients', async () => {
    server = await startBridgeServer({ userDataDir: dir, methods: { ping: () => 'pong' } })
    const a = connectClient(server.info.socketPath)
    const b = connectClient(server.info.socketPath)
    await a.ready
    await b.ready
    await handshake(a, server.info.token)
    await handshake(b, server.info.token)
    a.send({ protocol_version: 1, method: 'ping' })
    b.send({ protocol_version: 1, method: 'ping' })
    expect(JSON.parse(await a.next())).toEqual({ ok: true, result: 'pong' })
    expect(JSON.parse(await b.next())).toEqual({ ok: true, result: 'pong' })
    a.end()
    b.end()
  })

  it('stamps each connection with a distinct, stable client id', async () => {
    server = await startBridgeServer({
      userDataDir: dir,
      methods: { whoami: (_params, _signal, context) => ({ clientId: context.clientId }) },
    })
    const a = connectClient(server.info.socketPath)
    const b = connectClient(server.info.socketPath)
    await a.ready
    await b.ready
    await handshake(a, server.info.token)
    await handshake(b, server.info.token)
    a.send({ protocol_version: 1, method: 'whoami' })
    b.send({ protocol_version: 1, method: 'whoami' })
    expect(JSON.parse(await a.next())).toEqual({ ok: true, result: { clientId: 'conn-1' } })
    expect(JSON.parse(await b.next())).toEqual({ ok: true, result: { clientId: 'conn-2' } })
    // stable for the connection's lifetime (bridge turn ownership relies on it)
    a.send({ protocol_version: 1, method: 'whoami' })
    expect(JSON.parse(await a.next())).toEqual({ ok: true, result: { clientId: 'conn-1' } })
    a.end()
    b.end()
  })

  it('stop() removes the socket and the info file', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridgeServer({ userDataDir: dir, methods: {} })
    const socketPath = handle.info.socketPath
    await handle.stop()
    expect(existsSync(socketPath)).toBe(false)
    expect(existsSync(bridgeInfoPath(dir))).toBe(false)
    // a fresh server can bind the same socket again (stale socket cleanup)
    server = await startBridgeServer({ userDataDir: dir, methods: {} })
    expect(existsSync(server.info.socketPath)).toBe(true)
  })

  it('closes the connection when a peer pipelines past the queue cap', async () => {
    // A peer may only pipeline a bounded number of unanswered requests:
    // each one is a chained promise plus a buffered response, and a flood
    // from a client that never reads would grow both without bound.
    const stalls: (() => void)[] = []
    server = await startBridgeServer({
      userDataDir: dir,
      maxQueuedRequests: 4,
      methods: {
        stall: () =>
          new Promise((resolve) => {
            stalls.push(resolve)
          }),
      },
    })
    const client = connectClient(server.info.socketPath)
    await client.ready
    await handshake(client, server.info.token)
    const flood = `${Array.from({ length: 9 }, () => JSON.stringify({ protocol_version: 1, method: 'stall' })).join('\n')}\n`
    client.writeRaw(flood)
    // the rejection jumps the stalled FIFO chain on purpose: the connection
    // is being closed and the reason must reach the peer first
    const rejection = JSON.parse(await client.next())
    expect(rejection).toMatchObject({
      ok: false,
      error: { code: 'invalid_request', message: expect.stringContaining('pipelined') },
    })
    await client.closed
    for (const release of stalls) release()
    client.end()
  })

  it('does not dispatch lines written after the connection was closed', async () => {
    // BUG-901: close() half-closes the socket (end()), so data events keep
    // arriving while the error response flushes. Requests pipelined past a
    // queue-cap rejection (or an overflow close) used to run anyway with
    // their responses silently dropped — a retrying client would apply the
    // mutation twice, breaking the dispatcher's FIFO contract.
    const stalls: (() => void)[] = []
    let pings = 0
    server = await startBridgeServer({
      userDataDir: dir,
      maxQueuedRequests: 4,
      methods: {
        stall: () =>
          new Promise((resolve) => {
            stalls.push(resolve)
          }),
        ping: () => {
          pings += 1
          return 'pong'
        },
      },
    })
    const client = connectClient(server.info.socketPath)
    await client.ready
    await handshake(client, server.info.token)
    const flood = `${Array.from({ length: 9 }, () => JSON.stringify({ protocol_version: 1, method: 'stall' })).join('\n')}\n`
    client.writeRaw(flood)
    const rejection = JSON.parse(await client.next())
    expect(rejection).toMatchObject({
      ok: false,
      error: { code: 'invalid_request', message: expect.stringContaining('pipelined') },
    })
    // post-close lines must not reach the dispatcher
    client.writeRaw(`${JSON.stringify({ protocol_version: 1, method: 'ping' })}\n`)
    client.writeRaw(`${JSON.stringify({ protocol_version: 1, method: 'ping' })}\n`)
    await client.closed
    for (const release of stalls) release()
    client.end()
    expect(pings).toBe(0)
  })

  it('does not dispatch lines written after a line-overflow close', async () => {
    // BUG-901, overflow flavor: an oversized request closes the connection
    // the same way — anything the peer keeps writing afterwards is dropped
    // before dispatch, not executed into the void.
    let pings = 0
    server = await startBridgeServer({
      userDataDir: dir,
      methods: {
        ping: () => {
          pings += 1
          return 'pong'
        },
      },
    })
    const client = connectClient(server.info.socketPath)
    await client.ready
    await handshake(client, server.info.token)
    // one unterminated line past the default 8MB cap trips the overflow path
    client.writeRaw('x'.repeat(9 * 1024 * 1024))
    const rejection = JSON.parse(await client.next())
    expect(rejection).toMatchObject({
      ok: false,
      error: { code: 'invalid_request', message: expect.stringContaining('line limit') },
    })
    client.writeRaw(`${JSON.stringify({ protocol_version: 1, method: 'ping' })}\n`)
    await client.closed
    expect(pings).toBe(0)
  })
})

describe('createBackpressureWriter', () => {
  function fakeSocket() {
    const listeners: Record<string, (() => void)[]> = {}
    const calls: string[] = []
    const socket: BackpressureSocket & { calls: string[]; emit(event: string): void } = {
      calls,
      write: () => {
        calls.push('write')
        return calls.filter((call) => call === 'write').length < 2 // second write reports backpressure
      },
      pause: () => calls.push('pause'),
      resume: () => calls.push('resume'),
      on: (event, listener) => {
        ;(listeners[event] ??= []).push(listener)
      },
      emit: (event) => {
        for (const listener of listeners[event] ?? []) listener()
      },
    }
    return socket
  }

  it('pauses ingress when the socket buffer fills and resumes on drain', () => {
    const socket = fakeSocket()
    const writer = createBackpressureWriter(socket)
    expect(writer.paused).toBe(false)
    writer.write('first\n') // accepted without backpressure
    expect(writer.paused).toBe(false)
    writer.write('second\n') // write() now returns false
    expect(writer.paused).toBe(true)
    expect(socket.calls).toEqual(['write', 'write', 'pause'])
    socket.emit('drain')
    expect(writer.paused).toBe(false)
    expect(socket.calls).toEqual(['write', 'write', 'pause', 'resume'])
  })

  it('ignores drain events while flowing', () => {
    const socket = fakeSocket()
    const writer = createBackpressureWriter(socket)
    socket.emit('drain') // spurious drain with no pause outstanding
    expect(writer.paused).toBe(false)
    expect(socket.calls).toEqual([]) // no resume, no pause
  })
})

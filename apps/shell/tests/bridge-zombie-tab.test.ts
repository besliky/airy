import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { connect } from 'node:net'
import { EventEmitter } from 'node:events'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * BUG-1697: a killed docs renderer leaves a webContents that is NOT destroyed,
 * so 'destroyed' never fires and the bridge used to burn the full dispatcher
 * timeout on the corpse. These tests drive the REAL shell-bridge wiring (the
 * production methods map over a real unix-socket server, like
 * bridge-server.test.ts) against a fake TabManager/renderer and assert the
 * two healing guarantees: calls into a dead tab fail fast with tab_closed,
 * and a kill mid-call settles the pending call instead of hanging.
 */

const zombieIpcOn = vi.hoisted(() => vi.fn())

vi.mock('electron', () => ({ ipcMain: { on: zombieIpcOn } }))

import { BRIDGE_RESULT_CHANNEL } from '../../docs/src/shared/ipc'
import type { BridgeServerHandle } from '../src/main/bridge/server'
import { startShellBridge, stopShellBridge } from '../src/main/bridge/shell-bridge'
import type { TabManager } from '../src/main/tab-manager'

/** renderer stand-in: a real event emitter (once/off/emit) plus send/isDestroyed */
class FakeRendererWebContents extends EventEmitter {
  readonly id: number
  destroyed = false
  readonly sent: Array<{
    channel: string
    requestId: unknown
    method: unknown
    params: unknown
    clientId: unknown
  }> = []

  constructor(id: number) {
    super()
    this.id = id
  }

  isDestroyed(): boolean {
    return this.destroyed
  }

  send(channel: string, ...args: unknown[]): void {
    this.sent.push({
      channel,
      requestId: args[0],
      method: args[1],
      params: args[2],
      clientId: args[3],
    })
  }
}

/** the slice of TabManager the bridge reads (type-only import — no Electron) */
function managerStub(
  active: { webContents: FakeRendererWebContents; dead: boolean } | null,
): TabManager {
  return {
    activeDocsTab: () =>
      active ? { id: 't1', webContents: active.webContents, dead: active.dead } : undefined,
    docsTabSummaries: () =>
      active && !active.dead
        ? [{ id: 't1', title: 'r.docx', filePath: '/tmp/r.docx', active: true }]
        : [],
  } as unknown as TabManager
}

interface ClientLine {
  resolve(line: string): void
}

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

let dir: string
let server: BridgeServerHandle | null = null

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'airy-bridge-zombie-test-'))
})

afterEach(async () => {
  await stopShellBridge()
  server = null
  rmSync(dir, { recursive: true, force: true })
})

async function start(zombie: { webContents: FakeRendererWebContents; dead: boolean } | null) {
  // the real 30s dispatcher timeout: the assertions below only pass because
  // dead-tab calls return WITHOUT waiting it out
  server = await startShellBridge({
    userDataDir: dir,
    getTabManager: () => managerStub(zombie),
    timeoutMs: 30_000,
  })
  if (!server) throw new Error('bridge did not start')
  return server
}

/** reply from the fake renderer through the real BRIDGE_RESULT_CHANNEL handler */
function rendererReplies(wc: FakeRendererWebContents, requestId: unknown, result: unknown): void {
  const registration = zombieIpcOn.mock.calls.find(([channel]) => channel === BRIDGE_RESULT_CHANNEL)
  expect(registration).toBeDefined()
  const handler = registration![1] as (event: unknown, id: unknown, res: unknown) => void
  handler({ sender: { id: wc.id } }, requestId, result)
}

describe('bridge vs a killed docs renderer (BUG-1697)', () => {
  it('fails a call into a dead active tab fast with tab_closed', async () => {
    if (process.platform === 'win32') return
    const wc = new FakeRendererWebContents(101)
    const started = await start({ webContents: wc, dead: true })
    const client = connectClient(started.info.socketPath)
    await client.ready
    await handshake(client, started.info.token)

    const t0 = Date.now()
    client.send({ protocol_version: 1, method: 'get_context', params: {} })
    const reply = JSON.parse(await client.next())
    const elapsed = Date.now() - t0

    expect(reply).toMatchObject({ ok: false, error: { code: 'tab_closed' } })
    // fast-fail is single-digit ms; the pre-fix behavior was the full 30000ms
    expect(elapsed).toBeLessThan(2000)
    // the corpse renderer was never even asked
    expect(wc.sent).toHaveLength(0)
    client.end()
  })

  it('settles an in-flight call as tab_closed when the renderer is killed mid-call', async () => {
    if (process.platform === 'win32') return
    const wc = new FakeRendererWebContents(102)
    const started = await start({ webContents: wc, dead: false })
    const client = connectClient(started.info.socketPath)
    await client.ready
    await handshake(client, started.info.token)

    client.send({ protocol_version: 1, method: 'get_context', params: {} })
    await vi.waitFor(() => expect(wc.sent).toHaveLength(1))

    // kill -9: render-process-gone fires, the webContents is NOT destroyed
    const t0 = Date.now()
    wc.emit('render-process-gone', {}, { reason: 'killed' })
    const reply = JSON.parse(await client.next())
    const elapsed = Date.now() - t0

    expect(reply).toMatchObject({
      ok: false,
      error: { code: 'tab_closed', message: 'the document tab renderer is gone' },
    })
    expect(elapsed).toBeLessThan(2000)
    client.end()
  })

  it('healthy tabs still round-trip through the result channel', async () => {
    if (process.platform === 'win32') return
    const wc = new FakeRendererWebContents(103)
    const started = await start({ webContents: wc, dead: false })
    const client = connectClient(started.info.socketPath)
    await client.ready
    await handshake(client, started.info.token)

    client.send({ protocol_version: 1, method: 'get_context', params: {} })
    await vi.waitFor(() => expect(wc.sent).toHaveLength(1))
    expect(wc.sent[0]!.method).toBe('get_context')
    rendererReplies(wc, wc.sent[0]!.requestId, { ok: true, result: { blocks: [] } })

    expect(JSON.parse(await client.next())).toEqual({ ok: true, result: { blocks: [] } })
    client.end()
  })
})

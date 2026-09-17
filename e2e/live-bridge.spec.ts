import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createConnection } from 'node:net'
import { mkdtemp, copyFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Live-bridge round trip against the REAL running app: the shell's UDS bridge
 * server (the same one external MCP-style clients use) serves a hello/token
 * handshake, ping, and a get_context that must describe the open docs tab.
 * The MCP-server tests drive a mock bridge; this spec exercises the actual
 * socket, token file, and dispatcher end-to-end from a plain node:net client.
 */

interface BridgeInfo {
  socketPath: string
  token: string
  pid: number
  protocolVersion: number
}

/** minimal NDJSON client: one line per request, responses arrive in order */
class BridgeClient {
  private buffer = ''
  private pending: Array<{ resolve: (line: string) => void }> = []
  private lines: string[] = []
  constructor(private socket: import('node:net').Socket) {
    socket.on('data', (chunk: Buffer) => {
      this.buffer += chunk.toString('utf8')
      let index: number
      while ((index = this.buffer.indexOf('\n')) >= 0) {
        const line = this.buffer.slice(0, index)
        this.buffer = this.buffer.slice(index + 1)
        const waiter = this.pending.shift()
        if (waiter) waiter.resolve(line)
        else this.lines.push(line)
      }
    })
  }

  static async connect(socketPath: string): Promise<BridgeClient> {
    const socket = createConnection(socketPath)
    await new Promise<void>((resolvePromise, reject) => {
      socket.once('connect', resolvePromise)
      socket.once('error', reject)
    })
    return new BridgeClient(socket)
  }

  async call(method: string, params?: Record<string, unknown>): Promise<unknown> {
    const request: Record<string, unknown> = { protocol_version: 1, method }
    if (params) request.params = params
    this.socket.write(`${JSON.stringify(request)}\n`)
    const line = await new Promise<string>((resolvePromise) => {
      const buffered = this.lines.shift()
      if (buffered !== undefined) resolvePromise(buffered)
      else this.pending.push({ resolve: resolvePromise })
    })
    return JSON.parse(line)
  }

  end(): void {
    this.socket.end()
  }
}

test.describe('live bridge (UDS)', () => {
  test('hello handshake, ping, and get_context over the real socket', async () => {
    test.setTimeout(120_000)
    const scratch = await mkdtemp(join(tmpdir(), 'airy-bridge-e2e-'))
    const document = join(scratch, 'bridge-context.docx')
    await copyFile(resolve(__dirname, '../fixtures/generated/kitchen-sink.docx'), document)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'live-bridge',
      openFile: document,
    })
    try {
      // the docs tab must be live and active before get_context can serve it
      const editor = await waitForPageWithUrl(launched.app, 'docs/out')
      await editor.locator('.doc-page').first().waitFor({ timeout: 30_000 })

      // endpoint + token are published in the scratch userData dir this run owns
      const info = JSON.parse(
        await readFile(join(launched.userDataDir, 'airy-bridge.json'), 'utf8'),
      ) as BridgeInfo
      expect(info.protocolVersion).toBe(1)
      expect(info.token).toMatch(/^[0-9a-f]{64}$/)

      const client = await BridgeClient.connect(info.socketPath)
      try {
        const hello = (await client.call('hello', { token: info.token })) as {
          ok: boolean
          result: { server: string; protocolVersion: number }
        }
        expect(hello.ok).toBe(true)
        expect(hello.result.server).toBe('airy-bridge')
        expect(hello.result.protocolVersion).toBe(1)

        const ping = (await client.call('ping')) as {
          ok: boolean
          result: { pong: boolean }
        }
        expect(ping.ok).toBe(true)
        expect(ping.result.pong).toBe(true)

        // get_context serves the agent-facing document context: the open
        // file path plus a markdown description of the document's blocks
        const context = (await client.call('get_context')) as {
          ok: boolean
          result: { filePath: string; context: string }
        }
        expect(context.ok).toBe(true)
        expect(context.result.filePath).toBe(document)
        expect(typeof context.result.context).toBe('string')
        expect(context.result.context.length).toBeGreaterThan(0)

        // a wrong token on a fresh connection is rejected and the socket closed
        const intruder = await BridgeClient.connect(info.socketPath)
        const rejected = (await intruder.call('hello', {
          token: '0'.repeat(64),
        })) as { ok: boolean; error: { code: string } }
        expect(rejected.ok).toBe(false)
        expect(rejected.error.code).toBe('unauthorized')
        intruder.end()
      } finally {
        client.end()
      }
    } finally {
      await closeAndSaveVideo(launched, 'live-bridge')
    }
  })
})

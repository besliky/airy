// Integration for the live bridge tools: a real mock bridge server (UDS in a
// temp dir + 0600 info file, hello-gated FIFO NDJSON) serves the MCP client,
// which the tests drive through a linked InMemoryTransport pair — the same
// in-process pattern as server.test.ts / docx-tools.test.ts, so the full
// tool -> client -> socket -> dispatcher path runs per call.
import { existsSync, statSync } from 'node:fs'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, afterEach, beforeEach, describe, expect, it } from 'vitest'

import { buildServer } from '../src/index.js'
import { BRIDGE_INFO_FILE_ENV } from '../src/live/discovery.js'
import { BridgeClientError, createLiveBridge, resetSharedLiveBridge } from '../src/live/client.js'
import {
  MockBridgeMethodError,
  startMockBridge,
  type MockBridgeHandle,
} from './helpers/mock-bridge.js'

interface CallResult {
  isError?: boolean
  content?: Array<{ type: string; text?: string }>
  structuredContent?: Record<string, unknown>
}

interface Session {
  client: Client
  close: () => Promise<void>
}

async function connectSession(): Promise<Session> {
  const server = buildServer()
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'live-test-client', version: '0.0.1' })
  await Promise.all([client.connect(clientTransport), server.connect(serverTransport)])
  return {
    client,
    close: async () => {
      await Promise.all([client.close(), server.close()])
    },
  }
}

async function call(
  client: Client,
  name: string,
  args: Record<string, unknown>,
): Promise<CallResult> {
  return (await client.callTool({ name, arguments: args })) as CallResult
}

function text(result: CallResult): string {
  return result.content?.map((c) => c.text ?? '').join('') ?? ''
}

let dir: string
let bridge: MockBridgeHandle | null = null
let previousEnv: string | undefined

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'airy-live-test-'))
  previousEnv = process.env[BRIDGE_INFO_FILE_ENV]
  resetSharedLiveBridge()
})

afterEach(async () => {
  await bridge?.stop()
  bridge = null
  if (previousEnv === undefined) delete process.env[BRIDGE_INFO_FILE_ENV]
  else process.env[BRIDGE_INFO_FILE_ENV] = previousEnv
  resetSharedLiveBridge()
  rmSync(dir, { recursive: true, force: true })
})

afterAll(() => {
  resetSharedLiveBridge()
})

/** start a mock bridge and point the client's env at its info file */
async function startBridge(
  methods: Parameters<typeof startMockBridge>[0]['methods'],
  options: { timeoutMs?: number } = {},
): Promise<MockBridgeHandle> {
  bridge = await startMockBridge({ dir, methods, ...options })
  process.env[BRIDGE_INFO_FILE_ENV] = bridge.infoPath
  return bridge
}

const PING_METHODS = {
  ping: () => ({ pong: true, protocolVersion: 1, pid: 4242 }),
  list: () => ({
    documents: [{ id: 't1', title: 'Report.docx', filePath: '/docs/Report.docx', active: true }],
  }),
}

describe('live tools over MCP', () => {
  it('advertises the four live tools with annotations', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const byName = new Map(tools.map((t) => [t.name, t]))
      for (const name of ['live_status', 'live_get_context', 'live_apply_ops', 'live_undo']) {
        expect(byName.has(name)).toBe(true)
        // descriptions are written for external LLM agents (non-empty, English)
        expect((byName.get(name)?.description ?? '').length).toBeGreaterThan(40)
      }
      expect(byName.get('live_status')?.annotations?.readOnlyHint).toBe(true)
      expect(byName.get('live_get_context')?.annotations?.readOnlyHint).toBe(true)
      expect(byName.get('live_apply_ops')?.annotations?.destructiveHint).toBe(true)
      expect(byName.get('live_undo')?.annotations?.destructiveHint).toBe(true)
    } finally {
      await close()
    }
  })

  it('live_status reports a running bridge with its documents', async () => {
    if (process.platform === 'win32') return
    await startBridge(PING_METHODS)
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_status', {})
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent).toEqual({
        running: true,
        pid: 4242,
        protocolVersion: 1,
        documents: [
          { id: 't1', title: 'Report.docx', filePath: '/docs/Report.docx', active: true },
        ],
      })
      expect(text(result)).toContain('1 open document')
    } finally {
      await close()
    }
  })

  it('live_status returns {running:false} without an error when the bridge is off', async () => {
    process.env[BRIDGE_INFO_FILE_ENV] = join(dir, 'missing', 'airy-bridge.json')
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_status', {})
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent?.running).toBe(false)
      expect(String(result.structuredContent?.reason)).toContain('not running')
    } finally {
      await close()
    }
  })

  it('live_status reports not-running for a stale info file (dead socket)', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge(PING_METHODS)
    await handle.stop()
    bridge = null
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_status', {})
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent?.running).toBe(false)
    } finally {
      await close()
    }
  })

  it('live_get_context passes the active document context through', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge({
      get_context: () => ({
        context: { blocks: ['0|h1|Quarterly Report'], selection: '<sel>promising</sel>' },
        filePath: '/docs/Report.docx',
      }),
    })
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_get_context', {})
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent).toEqual({
        context: { blocks: ['0|h1|Quarterly Report'], selection: '<sel>promising</sel>' },
        filePath: '/docs/Report.docx',
      })
      expect(handle.requests.map((r) => r.method)).toEqual(['get_context'])
    } finally {
      await close()
    }
  })

  it('live_apply_ops inserts html then applies ops in one call', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge({
      insert_content: () => 'inserted 2 block(s)',
      apply_ops: () => 'setFont: changed 3 block(s)',
    })
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_apply_ops', {
        html: '<h2>Outlook</h2><p>Next quarter</p>',
        ops: [{ op: 'setFont', target: { nodeType: 'heading' }, italic: true }],
      })
      expect(result.isError).toBeFalsy()
      // insertion first so op targets can match the new content
      expect(handle.requests.map((r) => r.method)).toEqual(['insert_content', 'apply_ops'])
      expect(handle.requests[1]?.params.ops).toEqual([
        { op: 'setFont', target: { nodeType: 'heading' }, italic: true },
      ])
      expect(result.structuredContent).toEqual({
        insert: 'inserted 2 block(s)',
        ops: 'setFont: changed 3 block(s)',
      })
    } finally {
      await close()
    }
  })

  it('live_apply_ops surfaces a bridge error with its protocol code', async () => {
    if (process.platform === 'win32') return
    await startBridge({
      apply_ops: () => {
        throw new MockBridgeMethodError(
          'stale_document',
          'the document changed since the last get_context; refetch it',
        )
      },
    })
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_apply_ops', {
        ops: [{ op: 'setFont', target: { nodeType: 'heading' }, italic: true }],
      })
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('stale_document')
      expect(text(result)).toContain('refetch it')
    } finally {
      await close()
    }
  })

  it('live_apply_ops rejects a call with neither ops nor html', async () => {
    if (process.platform === 'win32') return
    await startBridge({})
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_apply_ops', {})
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('at least one of ops or html')
    } finally {
      await close()
    }
  })

  it('live_undo passes through and surfaces nothing_to_undo', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge({ undo: () => ({ undone: true }) })
    const { client, close } = await connectSession()
    try {
      const undone = await call(client, 'live_undo', {})
      expect(undone.isError).toBeFalsy()
      expect(undone.structuredContent).toEqual({ undone: true })
      expect(handle.requests.map((r) => r.method)).toEqual(['undo'])
    } finally {
      await close()
    }

    await handle.stop()
    bridge = await startBridge({
      undo: () => {
        throw new MockBridgeMethodError('nothing_to_undo', 'no bridge turn to undo yet')
      },
    })
    // drop the pooled connection so the next call reaches the new server
    resetSharedLiveBridge()
    const second = await connectSession()
    try {
      const result = await call(second.client, 'live_undo', {})
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('nothing_to_undo')
    } finally {
      await second.close()
    }
  })

  it('a wrong token is reported as unauthorized (and not-running by live_status)', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge({ get_context: () => ({ context: {}, filePath: null }) })
    // the client now reads a token the server will not accept
    await handle.rewriteInfo({ token: 'f'.repeat(64) })
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_get_context', {})
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('unauthorized')

      const status = await call(client, 'live_status', {})
      expect(status.isError).toBeFalsy()
      expect(status.structuredContent?.running).toBe(false)
      expect(String(status.structuredContent?.reason)).toContain('unauthorized')
    } finally {
      await close()
    }
  })

  it('a stalled bridge call times out with a clear error', async () => {
    if (process.platform === 'win32') return
    await startBridge(
      {
        ping: () => new Promise<never>(() => {}),
        list: () => new Promise<never>(() => {}),
      },
      { timeoutMs: 60 },
    )
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'live_status', {})
      // the stalled ping surfaces as not-running with the timeout as the reason
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent?.running).toBe(false)
      expect(String(result.structuredContent?.reason)).toContain('timeout')
    } finally {
      await close()
    }
  })
})

describe('live bridge client (direct)', () => {
  it('writes the info file with 0600 permissions', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge(PING_METHODS)
    expect(statSync(handle.infoPath).mode & 0o777).toBe(0o600)
    expect(existsSync(handle.socketPath)).toBe(true)
  })

  it('answers calls after the token handshake', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge(PING_METHODS)
    const bridgeClient = createLiveBridge()
    expect(await bridgeClient.call('ping')).toEqual({ pong: true, protocolVersion: 1, pid: 4242 })
    expect(handle.requests.map((r) => r.method)).toEqual(['ping'])
  })

  it('maps a missing info file to bridge_not_running listing the search', async () => {
    process.env[BRIDGE_INFO_FILE_ENV] = join(dir, 'nope.json')
    const bridgeClient = createLiveBridge()
    const err = await bridgeClient.call('ping').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(BridgeClientError)
    expect((err as BridgeClientError).code).toBe('bridge_not_running')
    expect((err as BridgeClientError).message).toContain(join(dir, 'nope.json'))
  })

  it('maps an unreachable socket to bridge_not_running', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge(PING_METHODS)
    // stop() removes the socket AND the info file; rewrite the info file so
    // discovery succeeds but the socket is gone (stale-file scenario)
    await handle.stop()
    await handle.rewriteInfo({})
    bridge = null
    const bridgeClient = createLiveBridge()
    const err = await bridgeClient.call('ping').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(BridgeClientError)
    expect((err as BridgeClientError).code).toBe('bridge_not_running')
    expect((err as BridgeClientError).message).toContain('cannot connect')
  })

  it('maps a rejected token to bridge_unauthorized', async () => {
    if (process.platform === 'win32') return
    const handle = await startBridge({})
    await handle.rewriteInfo({ token: '0'.repeat(64) })
    const bridgeClient = createLiveBridge()
    const err = await bridgeClient.call('ping').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(BridgeClientError)
    expect((err as BridgeClientError).code).toBe('bridge_unauthorized')
    expect((err as BridgeClientError).message).toContain('invalid bridge token')
  })

  it('passes server error codes through as bridge_error', async () => {
    if (process.platform === 'win32') return
    await startBridge({
      undo: () => {
        throw new MockBridgeMethodError('nothing_to_undo', 'no bridge turn to undo yet')
      },
    })
    const bridgeClient = createLiveBridge()
    const err = await bridgeClient.call('undo').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(BridgeClientError)
    expect((err as BridgeClientError).code).toBe('bridge_error')
    expect((err as BridgeClientError).bridgeCode).toBe('nothing_to_undo')
  })

  it('times a stalled call out server-side and keeps serving', async () => {
    if (process.platform === 'win32') return
    await startBridge(
      {
        stall: () => new Promise<never>(() => {}),
        ping: () => 'pong',
      },
      { timeoutMs: 60 },
    )
    const bridgeClient = createLiveBridge()
    const err = await bridgeClient.call('stall').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(BridgeClientError)
    expect((err as BridgeClientError).bridgeCode).toBe('timeout')
    // the server keeps answering after a timed-out call
    expect(await bridgeClient.call('ping')).toBe('pong')
  })

  it('times out client-side, drops the connection, and reconnects on the next call', async () => {
    if (process.platform === 'win32') return
    await startBridge(
      {
        swallow: () => new Promise((resolve) => setTimeout(() => resolve('late'), 400)),
        ping: () => 'pong',
      },
      { timeoutMs: 5_000 },
    )
    const bridgeClient = createLiveBridge({ timeoutMs: 80 })
    const err = await bridgeClient.call('swallow').catch((e: unknown) => e)
    expect(err).toBeInstanceOf(BridgeClientError)
    expect((err as BridgeClientError).bridgeCode).toBe('timeout')
    expect((err as BridgeClientError).message).toContain('timed out after 80ms')
    // no auto-reconnect on the dead stream, but the next call connects fresh
    expect(await bridgeClient.call('ping')).toBe('pong')
  })

  it('serializes concurrent calls strictly FIFO', async () => {
    if (process.platform === 'win32') return
    const settled: string[] = []
    await startBridge({
      slow: () => new Promise((resolve) => setTimeout(() => resolve('slow-result'), 60)),
      fast: () => 'fast-result',
    })
    const bridgeClient = createLiveBridge()
    const results = await Promise.all([
      bridgeClient.call('slow').then((value) => {
        settled.push('slow')
        return value
      }),
      bridgeClient.call('fast').then((value) => {
        settled.push('fast')
        return value
      }),
    ])
    expect(results).toEqual(['slow-result', 'fast-result'])
    // the client never overlaps requests: slow resolves first despite the delay
    expect(settled).toEqual(['slow', 'fast'])
  })
})

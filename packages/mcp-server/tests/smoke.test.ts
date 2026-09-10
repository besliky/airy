// Smoke test for the built bundle: spawns `node dist/index.js` and speaks raw
// JSON-RPC over stdio the way an MCP client host does. The package test script
// runs `npm run build` first, so dist/index.js exists by the time we get here.
import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { SERVER_NAME, SERVER_VERSION } from '../src/version.js'

const serverPath = fileURLToPath(new URL('../dist/index.js', import.meta.url))

// Legacy protocol revision (2024-11-05): the oldest stable version the SDK
// still negotiates, so the handshake is checked for backward compatibility.
const LEGACY_PROTOCOL_VERSION = '2024-11-05'

interface JsonRpcResponse {
  id?: number
  result?: {
    protocolVersion?: string
    serverInfo?: { name?: string; version?: string }
    tools?: Array<{ name: string }>
    structuredContent?: { pong?: boolean; server?: string }
    isError?: boolean
  }
  error?: { code: number; message: string }
}

interface Session {
  byId: Map<number, JsonRpcResponse>
  stderr: string
}

async function runStdioSession(): Promise<Session> {
  const child = spawn(process.execPath, [serverPath], { stdio: ['pipe', 'pipe', 'pipe'] })

  const responses: JsonRpcResponse[] = []
  const stderrChunks: string[] = []

  child.stdout.setEncoding('utf8')
  child.stderr.setEncoding('utf8')
  child.stdout.on('data', (chunk: string) => {
    for (const line of chunk.split('\n')) {
      const trimmed = line.trim()
      if (!trimmed) continue
      // stdout must carry nothing but valid JSON-RPC messages (stdio transport
      // spec) — a parse failure here fails the assertion below
      responses.push(JSON.parse(trimmed) as JsonRpcResponse)
    }
  })
  child.stderr.on('data', (chunk: string) => stderrChunks.push(chunk))

  const send = (message: unknown) => child.stdin.write(`${JSON.stringify(message)}\n`)

  send({
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: LEGACY_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: 'smoke-test', version: '0.0.1' },
    },
  })
  send({ jsonrpc: '2.0', method: 'notifications/initialized' })
  send({ jsonrpc: '2.0', id: 2, method: 'tools/list' })
  send({
    jsonrpc: '2.0',
    id: 3,
    method: 'tools/call',
    params: { name: 'ping', arguments: {} },
  })

  const withId = () => responses.filter((message) => typeof message.id === 'number')
  const deadline = Date.now() + 15000
  while (withId().length < 3) {
    if (Date.now() > deadline) {
      child.kill()
      throw new Error(`timed out waiting for responses, got ${JSON.stringify(responses)}`)
    }
    await new Promise((resolve) => setTimeout(resolve, 50))
  }

  child.kill()
  await new Promise<void>((resolve) => child.once('exit', () => resolve()))

  return {
    byId: new Map(withId().map((message) => [message.id as number, message])),
    stderr: stderrChunks.join(''),
  }
}

describe('dist bundle smoke test (stdio)', () => {
  it('answers initialize, tools/list and tools/call over stdio', async () => {
    const { byId, stderr } = await runStdioSession()

    const init = byId.get(1)
    expect(init?.error).toBeUndefined()
    expect(init?.result?.protocolVersion).toBe(LEGACY_PROTOCOL_VERSION)
    expect(init?.result?.serverInfo).toEqual({ name: SERVER_NAME, version: SERVER_VERSION })

    const list = byId.get(2)
    expect(list?.error).toBeUndefined()
    expect(list?.result?.tools?.map((tool) => tool.name)).toContain('ping')

    const call = byId.get(3)
    expect(call?.error).toBeUndefined()
    expect(call?.result?.isError).toBeFalsy()
    expect(call?.result?.structuredContent).toMatchObject({ pong: true, server: SERVER_NAME })

    // The startup notice goes to stderr, never to stdout
    expect(stderr).toContain(SERVER_NAME)
  })
})

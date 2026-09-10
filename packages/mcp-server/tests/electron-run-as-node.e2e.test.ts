// e2e: the installed app ships the MCP bundle in resources/mcp and agents run
// it with the app binary standing in for node (ELECTRON_RUN_AS_NODE=1, see
// apps/shell/electron-builder.cjs win/linux extraResources). This test proves
// that execution mode against the real electron binary from node_modules —
// not process.execPath — so a bundle that only works under plain node fails
// here instead of in a user's .mcp.json.
//
// Skipped when the electron package is not resolvable from the repo root
// (e.g. the test file copied elsewhere): outside a checkout there is no
// binary to stand in for the installed app.
import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { SERVER_NAME, SERVER_VERSION } from '../src/version.js'

const serverPath = fileURLToPath(new URL('../dist/index.js', import.meta.url))

// resolve 'electron' from the repo root, not from this package — the binary
// is hoisted to <repo>/node_modules/electron and require('electron') returns
// its executable path as a plain string (the package never launches anything)
const repoRootRequire = createRequire(
  fileURLToPath(new URL('../../../package.json', import.meta.url)),
)

let electronBinary: string | null = null
try {
  const resolved = repoRootRequire('electron') as unknown
  if (typeof resolved === 'string' && existsSync(resolved)) electronBinary = resolved
} catch {
  electronBinary = null
}

interface JsonRpcResponse {
  id?: number
  result?: {
    protocolVersion?: string
    serverInfo?: { name?: string; version?: string }
    tools?: Array<{ name: string }>
  }
  error?: { code: number; message: string }
}

describe('dist bundle under ELECTRON_RUN_AS_NODE (installed-app mode)', () => {
  it.skipIf(electronBinary === null)(
    'answers initialize and tools/list when run by the electron binary as node',
    async () => {
      const child = spawn(electronBinary as string, [serverPath], {
        stdio: ['pipe', 'pipe', 'pipe'],
        env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
      })

      const responses: JsonRpcResponse[] = []
      const stderrChunks: string[] = []
      child.stdout.setEncoding('utf8')
      child.stderr.setEncoding('utf8')
      child.stdout.on('data', (chunk: string) => {
        for (const line of chunk.split('\n')) {
          const trimmed = line.trim()
          // stdout carries nothing but JSON-RPC messages (stdio transport)
          if (trimmed) responses.push(JSON.parse(trimmed) as JsonRpcResponse)
        }
      })
      child.stderr.on('data', (chunk: string) => stderrChunks.push(chunk))

      const send = (message: unknown) => child.stdin.write(`${JSON.stringify(message)}\n`)
      send({
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: {
          protocolVersion: '2025-06-18',
          capabilities: {},
          clientInfo: { name: 'electron-run-as-node-e2e', version: '0.0.0' },
        },
      })
      send({ jsonrpc: '2.0', method: 'notifications/initialized' })
      send({ jsonrpc: '2.0', id: 2, method: 'tools/list' })

      const withId = () => responses.filter((m) => typeof m.id === 'number')
      const deadline = Date.now() + 15000
      while (withId().length < 2) {
        if (Date.now() > deadline) {
          child.kill()
          throw new Error(`timed out waiting for responses, got ${JSON.stringify(responses)}`)
        }
        await new Promise((resolve) => setTimeout(resolve, 50))
      }

      child.kill()
      await new Promise<void>((resolve) => child.once('exit', () => resolve()))

      const byId = new Map(withId().map((m) => [m.id as number, m]))
      const init = byId.get(1)
      expect(init?.error).toBeUndefined()
      expect(init?.result?.serverInfo).toEqual({ name: SERVER_NAME, version: SERVER_VERSION })

      const list = byId.get(2)
      expect(list?.error).toBeUndefined()
      expect(list?.result?.tools?.map((tool) => tool.name)).toContain('ping')
      expect(list?.result?.tools?.map((tool) => tool.name)).toContain('open_document')

      // the startup notice still goes to stderr, never to stdout
      expect(stderrChunks.join('')).toContain(SERVER_NAME)
    },
  )
})

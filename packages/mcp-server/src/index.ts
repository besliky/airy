#!/usr/bin/env node
// Airy Copilot MCP server entry point: serves MCP over stdio (JSON-RPC).
// stdout is reserved for protocol messages, so all logging goes to stderr
// (MCP stdio transport spec) — use console.error only, never console.log.
import { realpathSync } from 'node:fs'
import { pathToFileURL } from 'node:url'

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'

import { registerTools } from './tools/registry.js'
import { SERVER_NAME, SERVER_VERSION } from './version.js'

export { SERVER_NAME, SERVER_VERSION } from './version.js'

// Builds a configured server without attaching any transport, so tests can
// connect an in-memory transport pair instead of the process stdio.
export function buildServer(): McpServer {
  const server = new McpServer({
    name: SERVER_NAME,
    version: SERVER_VERSION,
  })
  registerTools(server)
  return server
}

// True when this module is the Node entry point (the esbuild bundle keeps
// import.meta.url, so this also works for dist/index.js and symlinked bins).
function isMainModule(): boolean {
  if (!process.argv[1]) {
    return false
  }
  try {
    return import.meta.url === pathToFileURL(realpathSync(process.argv[1])).href
  } catch {
    return false
  }
}

async function main(): Promise<void> {
  const server = buildServer()
  const transport = new StdioServerTransport()
  await server.connect(transport)
  console.error(`${SERVER_NAME} ${SERVER_VERSION} running on stdio`)
}

if (isMainModule()) {
  main().catch((err: unknown) => {
    console.error('Fatal:', err)
    process.exit(1)
  })
}

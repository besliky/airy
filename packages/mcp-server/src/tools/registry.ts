// Single registration point for every MCP tool (ADR-3): all registerTool calls
// live in this module, isolating the rest of the server from the SDK tool API
// so a future SDK migration (v1 -> v2) only has to touch this file.
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'

import { SERVER_NAME } from '../version.js'

export function registerTools(server: McpServer): void {
  // Liveness probe: lets an agent confirm the server is reachable before the
  // real document tools arrive in later phases.
  server.registerTool(
    'ping',
    {
      title: 'Ping',
      description: 'Check that the Airy MCP server is reachable and responsive',
      // Empty zod raw shape: the tool takes no arguments
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
      },
    },
    async () => {
      const structuredContent = {
        pong: true,
        server: SERVER_NAME,
        time: new Date().toISOString(),
      }
      // Structured output SHOULD also be mirrored as serialized JSON in a text
      // content block (MCP spec, Server Tools)
      return {
        content: [{ type: 'text', text: JSON.stringify(structuredContent) }],
        structuredContent,
      }
    },
  )
}

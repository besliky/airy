// In-process unit tests: a real Client talks to a fresh server over a linked
// InMemoryTransport pair — no process spawn, no stdio involved.
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { describe, expect, it } from 'vitest'

import { SERVER_NAME, SERVER_VERSION, buildServer } from '../src/index.js'

interface Session {
  client: Client
  close: () => Promise<void>
}

async function connectSession(): Promise<Session> {
  const server = buildServer()
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'test-client', version: '0.0.1' })
  await Promise.all([client.connect(clientTransport), server.connect(serverTransport)])
  return {
    client,
    close: async () => {
      await Promise.all([client.close(), server.close()])
    },
  }
}

describe('airy mcp server', () => {
  it('advertises its identity on initialize', async () => {
    const { client, close } = await connectSession()
    try {
      expect(client.getServerVersion()).toEqual({ name: SERVER_NAME, version: SERVER_VERSION })
    } finally {
      await close()
    }
  })

  it('lists the ping tool with a readOnlyHint annotation', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const ping = tools.find((tool) => tool.name === 'ping')
      expect(ping).toBeDefined()
      expect(ping?.title).toBe('Ping')
      expect(ping?.annotations?.readOnlyHint).toBe(true)
      // Empty zod raw shape compiles to a no-property object schema
      expect(ping?.inputSchema).toMatchObject({ type: 'object', properties: {} })
    } finally {
      await close()
    }
  })

  it('describes apply_ops with one shared line-ops guide, not per-format duplicates', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const applyOps = tools.find((tool) => tool.name === 'apply_ops')
      expect(applyOps?.description).toContain(
        'Markdown and html sessions accept a line-op vocabulary',
      )
      // the near-verbatim markdown/html guide pair used to double the text
      // in every tools/list answer
      expect(applyOps?.description).not.toContain('Markdown sessions accept')
      expect(applyOps?.description).not.toContain('HTML sessions accept')
    } finally {
      await close()
    }
  })

  it('discloses the docx save deltas in the save_document description', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const save = tools.find((tool) => tool.name === 'save_document')
      // OBS-1692: "untouched parts byte-identical" needs its docx exceptions
      // spelled out next to the promise, like the xlsx workbook.xml one —
      // core.xml is rewritten on purpose and the zip gains directory entries
      expect(save?.description).toContain('docProps/core.xml')
      expect(save?.description).toContain('cp:revision')
      expect(save?.description).toContain('dcterms:modified')
      expect(save?.description).toContain('directory entries')
      // the zero-edit guarantee stays intact next to the caveat
      expect(save?.description).toContain(
        'docx saves keep untouched parts byte-identical and a zero-edit save writes the original bytes back verbatim',
      )
    } finally {
      await close()
    }
  })

  it('returns structured content from ping', async () => {
    const { client, close } = await connectSession()
    try {
      const result = (await client.callTool({ name: 'ping', arguments: {} })) as {
        isError?: boolean
        structuredContent?: { pong: boolean; server: string; time: string }
        content?: Array<{ type: string; text?: string }>
      }
      expect(result.isError).toBeFalsy()
      const structured = result.structuredContent
      expect(structured?.pong).toBe(true)
      expect(structured?.server).toBe(SERVER_NAME)
      // time is an ISO 8601 timestamp
      expect(Number.isNaN(Date.parse(structured?.time ?? ''))).toBe(false)
      // The structured payload is mirrored as serialized JSON in text content
      expect(result.content).toHaveLength(1)
      expect(result.content?.[0]?.type).toBe('text')
      expect(JSON.parse(result.content?.[0]?.text ?? '')).toEqual(structured)
    } finally {
      await close()
    }
  })
})

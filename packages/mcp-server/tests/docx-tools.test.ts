// End-to-end MCP integration for the headless docx tools: a real Client talks
// to a fresh server over a linked InMemoryTransport pair (same pattern as
// server.test.ts) and drives the full open -> read -> insert -> ops -> save ->
// reopen cycle, plus the error paths an external agent will hit.
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'

import { buildServer } from '../src/index.js'
import { WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'
import { buildFixtureDocx } from './helpers/docx-fixture.js'

interface CallResult {
  isError?: boolean
  content?: Array<{ type: string; text?: string }>
  structuredContent?: Record<string, unknown>
}

interface Session {
  client: Client
  close: () => Promise<void>
}

let root: string
let previousRoot: string | undefined

async function connectSession(): Promise<Session> {
  const server = buildServer()
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'docx-test-client', version: '0.0.1' })
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

beforeAll(async () => {
  previousRoot = process.env[WORKSPACE_ROOT_ENV]
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-'))
  process.env[WORKSPACE_ROOT_ENV] = root
  await writeFile(join(root, 'report.docx'), await buildFixtureDocx())
})

afterAll(async () => {
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

// keep each test's side effects isolated: rewrite the fixture before every case
beforeEach(async () => {
  await writeFile(join(root, 'report.docx'), await buildFixtureDocx())
})

afterEach(async () => {
  await writeFile(join(root, 'report.docx'), await buildFixtureDocx())
})

describe('docx tools over MCP', () => {
  it('advertises the five document tools with annotations', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const names = tools.map((t) => t.name)
      for (const name of [
        'open_document',
        'read_document',
        'insert_content',
        'apply_ops',
        'save_document',
      ]) {
        expect(names).toContain(name)
      }
      const byName = new Map(tools.map((t) => [t.name, t]))
      expect(byName.get('open_document')?.annotations?.readOnlyHint).toBe(true)
      expect(byName.get('read_document')?.annotations?.readOnlyHint).toBe(true)
      expect(byName.get('save_document')?.annotations?.destructiveHint).toBe(true)
      // descriptions are written for external LLM agents (non-empty, English)
      for (const name of [
        'open_document',
        'read_document',
        'insert_content',
        'apply_ops',
        'save_document',
      ]) {
        expect((byName.get(name)?.description ?? '').length).toBeGreaterThan(40)
      }
    } finally {
      await close()
    }
  })

  it('runs the full edit cycle: open, read, insert, apply_ops(2+), save, reopen', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'report.docx' })
      expect(opened.isError).toBeFalsy()
      const handle = String(opened.structuredContent?.handle)
      expect(handle).toBeTruthy()
      expect(opened.structuredContent?.blockCount).toBe(7)
      // the tool returns the absolute path
      expect(opened.structuredContent?.path).toBe(join(root, 'report.docx'))

      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      expect(text(read)).toContain('0|h1|Quarterly Report')

      const inserted = await call(client, 'insert_content', {
        handle,
        html: '<h2>Outlook</h2><p>Next quarter looks <strong>promising</strong>.</p>',
        at: 1,
      })
      expect(inserted.isError).toBeFalsy()
      expect(inserted.structuredContent?.inserted).toBe(2)

      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [
          { op: 'findReplace', find: 'Revenue', replace: 'Profit' },
          {
            op: 'setParagraphFormat',
            target: { nodeType: 'heading', headingLevel: 2 },
            align: 'center',
          },
          { op: 'setFont', target: { blockIndexes: [3] }, italic: true },
        ],
      })
      expect(ops.isError).toBeFalsy()
      expect(String(ops.structuredContent?.summary)).toContain('findReplace: matched')

      const saved = await call(client, 'save_document', { handle, path: 'edited.docx' })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.path).toBe(join(root, 'edited.docx'))
      expect(saved.structuredContent?.unchanged).toBe(false)

      // reopen the saved copy and verify the changes landed
      const reopened = await call(client, 'open_document', { path: 'edited.docx' })
      expect(reopened.isError).toBeFalsy()
      const handle2 = String(reopened.structuredContent?.handle)
      expect(reopened.structuredContent?.blockCount).toBe(9)
      const reread = await call(client, 'read_document', { handle: handle2, blocks: [1, 2, 3] })
      const body = text(reread)
      expect(body).toContain('<p>Profit grew by <strong>12 percent</strong> year over year.</p>')
      expect(body).toContain('<h2>Outlook</h2>')
      // setFont italic wrapped the whole block; the bold mark nests inside
      expect(body).toContain('<em><strong>promising</strong></em>')
      expect(body).toMatch(/2\|h2\|Outlook/)
    } finally {
      await close()
    }
  })

  it('saves in place byte-identically when nothing was edited', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'report.docx' })
      const handle = String(opened.structuredContent?.handle)
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.unchanged).toBe(true)
      expect(
        Buffer.compare(await readFile(join(root, 'report.docx')), await buildFixtureDocx()),
      ).toBe(0)
    } finally {
      await close()
    }
  })

  it('reports unknown handles as tool errors', async () => {
    const { client, close } = await connectSession()
    try {
      const result = await call(client, 'read_document', { handle: 'no-such-handle' })
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('Unknown document handle')
    } finally {
      await close()
    }
  })

  it('rejects an invalid ops batch atomically with a clear message', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'report.docx' })
      const handle = String(opened.structuredContent?.handle)
      const bad = await call(client, 'apply_ops', {
        handle,
        ops: [
          { op: 'findReplace', find: 'Revenue', replace: 'Profit' },
          { op: 'setFont', target: { nodeType: 'paragraph' }, bogus: 1 },
        ],
      })
      expect(bad.isError).toBe(true)
      expect(text(bad)).toContain('unknown field')
      expect(text(bad)).toContain('nothing was applied')
      // the valid first op must not have landed: a save is byte-identical
      const saved = await call(client, 'save_document', { handle, path: 'atomic.docx' })
      expect(saved.structuredContent?.unchanged).toBe(true)
    } finally {
      await close()
    }
  })

  it('rejects paths outside the workspace root', async () => {
    const { client, close } = await connectSession()
    try {
      const outside = await call(client, 'open_document', { path: '../../etc/passwd' })
      expect(outside.isError).toBe(true)
      expect(text(outside)).toContain('outside the workspace root')
      const abs = await call(client, 'open_document', { path: '/etc/hosts' })
      expect(abs.isError).toBe(true)
      expect(text(abs)).toContain('outside the workspace root')
    } finally {
      await close()
    }
  })

  it('refuses to save when the file changed on disk since open', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'report.docx' })
      const handle = String(opened.structuredContent?.handle)
      await call(client, 'insert_content', { handle, html: '<p>x</p>', at: 0 })
      // external writer flips a byte (same size, mtime moves)
      const raw = new Uint8Array(await readFile(join(root, 'report.docx')))
      raw[raw.length - 1] ^= 0xff
      await writeFile(join(root, 'report.docx'), raw)
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBe(true)
      expect(text(saved)).toContain('changed on disk')
    } finally {
      await close()
    }
  })

  it('opens relative to the workspace root and reports missing files clearly', async () => {
    const { client, close } = await connectSession()
    try {
      const missing = await call(client, 'open_document', { path: 'nope.docx' })
      expect(missing.isError).toBe(true)
      expect(text(missing)).toContain('Cannot read')
    } finally {
      await close()
    }
  })
})

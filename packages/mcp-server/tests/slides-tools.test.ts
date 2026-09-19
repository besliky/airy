// MCP-level coverage for the slides tools (PAR-001): a real Client over an
// InMemoryTransport pair drives open_document/read_deck/insert_content/
// save_document/close_document over a generated .pptx fixture.
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, beforeAll, beforeEach, describe, expect, it } from 'vitest'

import { buildServer } from '../src/index.js'
import { WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'
import { buildFixturePptx } from './helpers/pptx-fixture.js'

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
  const client = new Client({ name: 'slides-test-client', version: '0.0.1' })
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
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-slides-'))
  process.env[WORKSPACE_ROOT_ENV] = root
})

afterAll(async () => {
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

beforeEach(async () => {
  await writeFile(join(root, 'deck.pptx'), await buildFixturePptx())
})

describe('slides tools over MCP', () => {
  it('advertises read_deck with a readOnlyHint annotation', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const readDeck = tools.find((tool) => tool.name === 'read_deck')
      expect(readDeck).toBeDefined()
      expect(readDeck?.title).toBe('Read presentation')
      expect(readDeck?.annotations?.readOnlyHint).toBe(true)
      expect((readDeck?.description ?? '').length).toBeGreaterThan(40)
    } finally {
      await close()
    }
  })

  it('runs the slides cycle: open, deck overview, slide detail, insert, save, close', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'deck.pptx' })
      expect(opened.isError).toBeFalsy()
      expect(opened.structuredContent?.kind).toBe('slides')
      expect(opened.structuredContent?.editable).toBe(true)
      expect(opened.structuredContent?.slideCount).toBe(1)
      expect(text(opened)).toContain('editable slides session')
      const handle = String(opened.structuredContent?.handle)

      const overview = await call(client, 'read_deck', { handle })
      expect(overview.isError).toBeFalsy()
      expect(text(overview)).toContain('The deck has 1 slide(s)')
      expect(text(overview)).toMatch(/^0\|2\|Quarterly Review$/m)

      const detail = await call(client, 'read_deck', { handle, slide: 0 })
      expect(text(detail)).toContain('Slide 0 of 1')
      expect(text(detail)).toMatch(/^0\|text\|TextBox 2\|Quarterly Review$/m)

      const inserted = await call(client, 'insert_content', {
        handle,
        slide: 0,
        text: 'Added over MCP',
        x: 2,
        y: 2,
        width: 5,
        height: 0.8,
      })
      expect(inserted.isError).toBeFalsy()
      expect(inserted.structuredContent?.added).toBe(true)
      expect(inserted.structuredContent?.element).toBe(2)
      expect(inserted.structuredContent?.dirty).toBe(true)

      const replaced = await call(client, 'insert_content', {
        handle,
        slide: 0,
        slideElement: 1,
        text: 'Rewritten over MCP',
      })
      expect(replaced.isError).toBeFalsy()
      expect(replaced.structuredContent?.added).toBe(false)

      const saved = await call(client, 'save_document', { handle, path: 'copy.pptx' })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.path).toBe(join(root, 'copy.pptx'))
      expect(saved.structuredContent?.format).toBe('pptx')
      expect(saved.structuredContent?.unchanged).toBe(false)
      // the saved deck reopens as a valid pptx with both edits
      const reopened = await call(client, 'open_document', { path: 'copy.pptx' })
      const reopenedHandle = String(reopened.structuredContent?.handle)
      expect(reopened.structuredContent?.elementCount).toBe(3)
      const slide = await call(client, 'read_deck', { handle: reopenedHandle, slide: 0 })
      expect(text(slide)).toContain('Added over MCP')
      expect(text(slide)).toContain('Rewritten over MCP')

      const closed = await call(client, 'close_document', { handle })
      expect(closed.structuredContent?.closed).toBe(true)
      const after = await call(client, 'read_deck', { handle })
      expect(after.isError).toBe(true)
      expect(text(after)).toContain('Unknown document handle')
      await call(client, 'close_document', { handle: reopenedHandle })
    } finally {
      await close()
    }
  })

  it('routes tools by session kind with actionable errors', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'deck.pptx' })
      const handle = String(opened.structuredContent?.handle)

      const wrongRead = await call(client, 'read_document', { handle })
      expect(wrongRead.isError).toBe(true)
      expect(text(wrongRead)).toContain('use read_deck')

      const wrongOps = await call(client, 'apply_ops', { handle, ops: [{ op: 'nope' }] })
      expect(wrongOps.isError).toBe(true)
      expect(text(wrongOps)).toContain('only available for editable .docx sessions')

      const wrongInsert = await call(client, 'insert_content', {
        handle,
        slide: 0,
        html: '<p>no html here</p>',
      })
      expect(wrongInsert.isError).toBe(true)
      expect(text(wrongInsert)).toContain('plain text in `text`')

      const missingSlide = await call(client, 'insert_content', {
        handle,
        text: 'where?',
      })
      expect(missingSlide.isError).toBe(true)
      expect(text(missingSlide)).toContain('needs `slide`')

      const badSlide = await call(client, 'read_deck', { handle, slide: 5 })
      expect(badSlide.isError).toBe(true)
      expect(text(badSlide)).toContain('slide index 5 is out of range')

      const wrongFormat = await call(client, 'save_document', { handle, format: 'docx' })
      expect(wrongFormat.isError).toBe(true)
      expect(text(wrongFormat)).toContain('not valid for a slides session')

      const docxOpts = await call(client, 'insert_content', {
        handle,
        text: 'x',
        marker: 'nope',
        slide: 0,
      })
      expect(docxOpts.isError).toBe(true)
      expect(text(docxOpts)).toContain('positions inserts via slide')
      await call(client, 'close_document', { handle })
    } finally {
      await close()
    }
  })

  it('refuses legacy .ppt with a conversion hint and keeps unsupported errors', async () => {
    const { client, close } = await connectSession()
    try {
      await writeFile(join(root, 'legacy.ppt'), 'not a real ppt')
      const legacy = await call(client, 'open_document', { path: 'legacy.ppt' })
      expect(legacy.isError).toBe(true)
      expect(text(legacy)).toContain('Convert the deck to .pptx')

      const bad = await call(client, 'open_document', { path: 'archive.zip' })
      expect(bad.isError).toBe(true)
      expect(text(bad)).toContain('Unsupported file type ".zip"')

      const outside = await call(client, 'open_document', { path: '../../deck.pptx' })
      expect(outside.isError).toBe(true)
      expect(text(outside)).toContain('outside the workspace root')
    } finally {
      await close()
    }
  })

  it('refuses to clobber an existing save target unless overwrite is passed', async () => {
    const { client, close } = await connectSession()
    try {
      await writeFile(join(root, 'precious.pptx'), 'unrelated deck')
      const opened = await call(client, 'open_document', { path: 'deck.pptx' })
      const handle = String(opened.structuredContent?.handle)
      await call(client, 'insert_content', { handle, slide: 0, text: 'Edit' })

      const refused = await call(client, 'save_document', { handle, path: 'precious.pptx' })
      expect(refused.isError).toBe(true)
      expect(text(refused)).toContain('overwrite: true')
      expect(await readFile(join(root, 'precious.pptx'), 'utf8')).toBe('unrelated deck')

      const forced = await call(client, 'save_document', {
        handle,
        path: 'precious.pptx',
        overwrite: true,
      })
      expect(forced.isError).toBeFalsy()
      await call(client, 'close_document', { handle })
    } finally {
      await close()
    }
  })
})

// MCP-level coverage for the S6 multi-format tools: a real Client over an
// InMemoryTransport pair drives open_document/read_document/read_workbook/
// save_document/close_document across session kinds, with the sidecar IO
// stubbed and the save gateway mocked (the real gateway needs the sidecar
// binary; see xlsx-integration.test.ts). LibreOffice is mocked absent so the
// .doc read-only fallback is deterministic.
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('../src/xlsx/save.js', () => ({
  saveWorkbookViaSidecar: vi.fn(async (request: { targetPath: string }) => {
    await writeFile(request.targetPath, 'saved-xlsx-bytes')
    return { touchedEntries: [], removedEntries: [], addedEntries: [] }
  }),
}))

vi.mock('../src/import/soffice.js', () => ({
  SOFFICE_FILTERS: { docx: 'MS Word 2007 XML', doc: 'MS Word 97', odt: 'writer8', ods: 'calc8' },
  findSoffice: vi.fn(async () => null),
  sofficeMissingError: (reason: string) => new Error(`${reason} Install LibreOffice.`),
  convertViaSoffice: vi.fn(async () => {
    throw new Error('not available in unit tests')
  }),
}))

import { buildServer } from '../src/index.js'
import { WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'
import { setSharedIo } from '../src/xlsx/session.js'
import { makeStubIo } from './helpers/stub-sidecar.js'
import { buildFixtureDocx } from './helpers/docx-fixture.js'

const DOC_SAMPLE = fileURLToPath(
  new URL('../../../packages/file-parse/tests/fixtures/legacy-sample.doc', import.meta.url),
)

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
  const client = new Client({ name: 'multi-format-test-client', version: '0.0.1' })
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
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-multi-'))
  process.env[WORKSPACE_ROOT_ENV] = root
})

afterAll(async () => {
  setSharedIo(null)
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

beforeEach(async () => {
  await writeFile(join(root, 'book.xlsx'), 'stub-xlsx-bytes')
  await writeFile(join(root, 'report.docx'), await buildFixtureDocx())
  const { copyFile } = await import('node:fs/promises')
  await copyFile(DOC_SAMPLE, join(root, 'legacy.doc'))
  setSharedIo(
    makeStubIo({
      cells: [
        { row: 0, column: 0, value: 'Region' },
        { row: 0, column: 1, value: 'Sales' },
        { row: 1, column: 1, value: 42 },
      ],
    }),
  )
})

describe('multi-format document tools over MCP', () => {
  it('advertises the workbook + close tools with annotations', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const names = tools.map((t) => t.name)
      for (const name of [
        'open_document',
        'read_document',
        'read_workbook',
        'save_document',
        'close_document',
      ]) {
        expect(names).toContain(name)
      }
      const byName = new Map(tools.map((t) => [t.name, t]))
      expect(byName.get('read_workbook')?.annotations?.readOnlyHint).toBe(true)
      expect(byName.get('close_document')?.annotations?.destructiveHint).toBe(true)
      for (const name of ['open_document', 'read_workbook', 'close_document']) {
        expect((byName.get(name)?.description ?? '').length).toBeGreaterThan(40)
      }
    } finally {
      await close()
    }
  })

  it('runs the xlsx cycle: open, overview, range read, save, close', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'book.xlsx' })
      expect(opened.isError).toBeFalsy()
      expect(opened.structuredContent?.kind).toBe('xlsx')
      expect(opened.structuredContent?.editable).toBe(true)
      expect(opened.structuredContent?.converted).toBe(false)
      expect(String(opened.structuredContent?.handle)).toBeTruthy()
      const handle = String(opened.structuredContent?.handle)
      expect(text(opened)).toContain('editable workbook')

      const overview = await call(client, 'read_workbook', { handle })
      expect(overview.isError).toBeFalsy()
      expect(text(overview)).toMatch(/0\|Sheet1\|sheet-0\|100 x 26/)

      const range = await call(client, 'read_workbook', { handle, sheet: 'Sheet1', range: 'A1:B2' })
      expect(range.isError).toBeFalsy()
      expect(text(range)).toContain('|1|Region|Sales|')
      expect(text(range)).toContain('|2||42|')

      const saved = await call(client, 'save_document', { handle, path: 'copy.xlsx' })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.path).toBe(join(root, 'copy.xlsx'))
      expect(saved.structuredContent?.format).toBe('xlsx')

      const closed = await call(client, 'close_document', { handle })
      expect(closed.isError).toBeFalsy()
      expect(closed.structuredContent?.closed).toBe(true)
      // the handle is gone after close
      const after = await call(client, 'read_workbook', { handle })
      expect(after.isError).toBe(true)
      expect(text(after)).toContain('Unknown document handle')
    } finally {
      await close()
    }
  })

  it('routes read_document vs read_workbook by session kind', async () => {
    const { client, close } = await connectSession()
    try {
      const docx = await call(client, 'open_document', { path: 'report.docx' })
      const docxHandle = String(docx.structuredContent?.handle)
      const wrongTool = await call(client, 'read_workbook', { handle: docxHandle })
      expect(wrongTool.isError).toBe(true)
      expect(text(wrongTool)).toContain('not a workbook session')

      const book = await call(client, 'open_document', { path: 'book.xlsx' })
      const bookHandle = String(book.structuredContent?.handle)
      const wrongTool2 = await call(client, 'read_document', { handle: bookHandle })
      expect(wrongTool2.isError).toBe(true)
      expect(text(wrongTool2)).toContain('use read_workbook')
      await call(client, 'close_document', { handle: bookHandle })
    } finally {
      await close()
    }
  })

  it('opens legacy .doc read-only without LibreOffice and refuses to save it', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'legacy.doc' })
      expect(opened.isError).toBeFalsy()
      expect(opened.structuredContent?.kind).toBe('text')
      expect(opened.structuredContent?.editable).toBe(false)
      expect(opened.structuredContent?.format).toBe('doc')
      const handle = String(opened.structuredContent?.handle)

      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      expect(text(read)).toContain('Legacy Report')
      expect(text(read)).toContain('editable: false')

      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBe(true)
      expect(text(saved)).toContain('read-only text session')
    } finally {
      await close()
    }
  })

  it('rejects unsupported formats and format/session mismatches', async () => {
    const { client, close } = await connectSession()
    try {
      const bad = await call(client, 'open_document', { path: 'deck.pptx' })
      expect(bad.isError).toBe(true)
      expect(text(bad)).toContain('Unsupported file type ".pptx"')

      const book = await call(client, 'open_document', { path: 'book.xlsx' })
      const bookHandle = String(book.structuredContent?.handle)
      const wrongFormat = await call(client, 'save_document', {
        handle: bookHandle,
        format: 'docx',
      })
      expect(wrongFormat.isError).toBe(true)
      expect(text(wrongFormat)).toContain('not valid for a workbook session')
      await call(client, 'close_document', { handle: bookHandle })

      const doc = await call(client, 'open_document', { path: 'report.docx' })
      const docHandle = String(doc.structuredContent?.handle)
      const wrongFormat2 = await call(client, 'save_document', {
        handle: docHandle,
        format: 'origin',
      })
      expect(wrongFormat2.isError).toBe(true)
      expect(text(wrongFormat2)).toContain('only valid for sessions converted from .doc/.odt')
    } finally {
      await close()
    }
  })
})

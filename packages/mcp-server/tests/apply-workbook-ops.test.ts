// MCP-level coverage for apply_workbook_ops: a real Client over an
// InMemoryTransport pair drives the tool across the happy path (value +
// formula + style), dryRun, wrong-handle rejection, batch-cap enforcement and
// the invalid-batch atomicity — with the sidecar IO stubbed and the save
// gateway mocked (the real gateway needs the sidecar binary; covered by
// xlsx-integration.test.ts).
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

// capture of every saveWorkbookViaSidecar call; the mock writes the target so
// the session's post-save stat works
const saveCalls: Array<Record<string, unknown>> = []
vi.mock('../src/xlsx/save.js', async (importOriginal) => {
  const original = await importOriginal<typeof import('../src/xlsx/save.js')>()
  return {
    ...original,
    saveWorkbookViaSidecar: vi.fn(async (request: Record<string, unknown>) => {
      saveCalls.push(request)
      await writeFile(String(request.targetPath), 'saved-xlsx-bytes')
      return { touchedEntries: ['xl/worksheets/sheet1.xml'], removedEntries: [], addedEntries: [] }
    }),
  }
})

import { buildServer } from '../src/index.js'
import { WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'
import { setSharedIo } from '../src/xlsx/session.js'
import { makeStubIo } from './helpers/stub-sidecar.js'
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
  const client = new Client({ name: 'workbook-ops-test-client', version: '0.0.1' })
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

/** open book.xlsx through the MCP tool and return the workbook handle */
async function openBook(client: Client): Promise<string> {
  const opened = await call(client, 'open_document', { path: 'book.xlsx' })
  expect(opened.isError).toBeFalsy()
  return String(opened.structuredContent?.handle)
}

beforeAll(async () => {
  previousRoot = process.env[WORKSPACE_ROOT_ENV]
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-ops-'))
  process.env[WORKSPACE_ROOT_ENV] = root
})

afterAll(async () => {
  setSharedIo(null)
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

beforeEach(async () => {
  saveCalls.length = 0
  await writeFile(join(root, 'book.xlsx'), 'stub-xlsx-bytes')
  await writeFile(join(root, 'report.docx'), await buildFixtureDocx())
  setSharedIo(makeStubIo())
})

describe('apply_workbook_ops over MCP', () => {
  it('advertises the tool with a non-destructive annotation', async () => {
    const { client, close } = await connectSession()
    try {
      const { tools } = await client.listTools()
      const tool = tools.find((t) => t.name === 'apply_workbook_ops')
      expect(tool).toBeTruthy()
      expect((tool?.description ?? '').length).toBeGreaterThan(40)
      expect(tool?.annotations?.readOnlyHint).toBeUndefined()
      expect(tool?.annotations?.destructiveHint).toBe(false)
    } finally {
      await close()
    }
  })

  it('journeys value + formula + style edits into the save pipeline', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const result = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [
          { sheet: 'Sheet1', ref: 'A1', value: 'Region' },
          { sheet: 'Sheet1', ref: 'B2', formula: 'SUM(A1:A1)' },
          { sheet: 'Data', ref: 'C3', style: { bold: true, fillColor: '#FFEE00' } },
          { sheet: 1, ref: 'D4', value: 7, style: { numberFormat: '0.00' } },
        ],
      })
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent).toEqual({ journaled: 4, dirty: true, dryRun: false })
      expect(text(result)).toContain('persist with save_document')

      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saveCalls).toHaveLength(1)
      const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
      expect(edits).toHaveLength(4)
      expect(edits[0]).toMatchObject({
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: 'Region' },
      })
      expect(edits[1]).toMatchObject({
        sheetName: 'Sheet1',
        row: 1,
        column: 1,
        writeValue: true,
        cell: { value: '', formula: '=SUM(A1:A1)' },
      })
      // style-only edit keeps the stored content untouched
      expect(edits[2]).toMatchObject({
        sheetName: 'Data',
        row: 2,
        column: 2,
        writeValue: false,
        style: { bold: true, fillColor: '#FFEE00' },
      })
      // value + style combine into one journaled edit
      expect(edits[3]).toMatchObject({
        sheetName: 'Data',
        row: 3,
        column: 3,
        writeValue: true,
        cell: { value: 7 },
        style: { numberFormat: '0.00' },
      })
    } finally {
      await close()
    }
  })

  it('accepts minimal rich runs: the style flags are optional (absent = plain)', async () => {
    // the schema used to require all four booleans per run, so an honest
    // minimal call failed zod with "expected boolean, received undefined at
    // italic" — the flags serialize truthily and must be optional
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const result = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [
          {
            sheet: 'Sheet1',
            ref: 'A1',
            rich: [{ text: 'Big ', bold: true }, { text: 'news' }],
          },
        ],
      })
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent).toEqual({ journaled: 1, dirty: true, dryRun: false })

      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
      expect(edits[0]).toMatchObject({
        writeValue: true,
        cell: { value: 'Big news' },
        rich: [{ text: 'Big ', bold: true }, { text: 'news' }],
      })
    } finally {
      await close()
    }
  })

  it('reports the merged journal count, not the raw input edit count', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const result = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [
          { sheet: 'Sheet1', ref: 'A1', value: 'one' },
          { sheet: 'Sheet1', ref: 'A1', value: 'two' },
          { sheet: 'Sheet1', ref: 'A1', style: { bold: true } },
          { sheet: 'Sheet1', ref: 'A1', style: { italic: true } },
        ],
      })
      expect(result.isError).toBeFalsy()
      // four edits to one cell collapse into a single journaled entry
      expect(result.structuredContent).toEqual({ journaled: 1, dirty: true, dryRun: false })
      expect(text(result)).toContain('Journaled 1 cell edit(s)')
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saveCalls[0]?.edits).toHaveLength(1)
    } finally {
      await close()
    }
  })

  it('dryRun validates and reports without journaling anything', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const result = await call(client, 'apply_workbook_ops', {
        handle,
        dryRun: true,
        edits: [{ sheet: 'Sheet1', ref: 'A1', value: 'kept for the dry run' }],
      })
      expect(result.isError).toBeFalsy()
      expect(result.structuredContent).toEqual({ journaled: 1, dirty: false, dryRun: true })

      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saveCalls[0]?.edits).toEqual([])
      expect(saved.structuredContent?.unchanged).toBe(true)
    } finally {
      await close()
    }
  })

  it('rejects docx handles with a pointed error', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'report.docx' })
      const docxHandle = String(opened.structuredContent?.handle)
      const result = await call(client, 'apply_workbook_ops', {
        handle: docxHandle,
        edits: [{ sheet: 'Sheet1', ref: 'A1', value: 1 }],
      })
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('only available for workbook sessions')
      expect(text(result)).toContain('"docx"')
    } finally {
      await close()
    }
  })

  it('enforces the 100-edit batch cap', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const edits = Array.from({ length: 101 }, (_, i) => ({
        sheet: 'Sheet1',
        ref: `A${String(i + 1)}`,
        value: i,
      }))
      const result = await call(client, 'apply_workbook_ops', { handle, edits })
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('100')
      // nothing was journaled by the rejected batch
      const saved = await call(client, 'save_document', { handle })
      expect(saved.structuredContent?.unchanged).toBe(true)
    } finally {
      await close()
    }
  })

  it('validates the whole batch up front: a bad edit journals nothing', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const result = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [
          { sheet: 'Sheet1', ref: 'A1', value: 'would be journaled' },
          { sheet: 'Nope', ref: 'A1', value: 'unknown sheet' },
        ],
      })
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('No sheet "Nope"')

      const rangeRef = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [
          { sheet: 'Sheet1', ref: 'A1', value: 'kept' },
          { sheet: 'Sheet1', ref: 'A1:B2', value: 'range refs are rejected' },
        ],
      })
      expect(rangeRef.isError).toBe(true)
      expect(text(rangeRef)).toContain('must be a single cell')

      const empty = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [{ sheet: 'Sheet1', ref: 'A1' }],
      })
      expect(empty.isError).toBe(true)
      expect(text(empty)).toContain('at least one of value, formula')

      // none of the rejected batches journaled anything
      const saved = await call(client, 'save_document', { handle })
      expect(saved.structuredContent?.unchanged).toBe(true)
      expect(saveCalls[0]?.edits).toEqual([])
    } finally {
      await close()
    }
  })

  it('rejects row-less refs ("A0") at the op; the session survives and saves (BUG-1632)', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openBook(client)
      const result = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [{ sheet: 'Sheet1', ref: 'A0', value: 1 }],
      })
      expect(result.isError).toBe(true)
      expect(text(result)).toContain('addresses no cell')

      // the audit repro: this op used to answer "Journaled 1 cell edit(s)"
      // with row -1 in the journal, and every later save then died with
      // "Invalid cell coordinates: -1,0" until close lost the unsaved edits.
      // The refused op leaves the journal clean: the next save works.
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.unchanged).toBe(true)
      expect(saveCalls[0]?.edits).toEqual([])

      // and the session is still fully usable for valid edits
      const retry = await call(client, 'apply_workbook_ops', {
        handle,
        edits: [{ sheet: 'Sheet1', ref: 'A1', value: 'works again' }],
      })
      expect(retry.isError).toBeFalsy()
      const savedAgain = await call(client, 'save_document', { handle })
      expect(savedAgain.isError).toBeFalsy()
      expect(saveCalls[1]?.edits).toHaveLength(1)
    } finally {
      await close()
    }
  })
})

describe('apply_workbook_ops journal semantics (session level)', () => {
  it('merges later edits to the same cell per channel, content last-wins', async () => {
    const { XlsxSession } = await import('../src/xlsx/session.js')
    const session = await XlsxSession.open(join(root, 'book.xlsx'), root, makeStubIo())
    session.setCells({
      sheet: 'Sheet1',
      cells: [
        { ref: 'A1', value: 'first' },
        { ref: 'A1', value: 'second' },
        { ref: 'A1', style: { bold: true } },
        { ref: 'B2', formula: 'SUM(A1:A1)' },
      ],
    })
    await session.save()
    const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
    expect(edits).toHaveLength(2)
    // content replaced by the later edit; the style patch merged on top
    expect(edits[0]).toMatchObject({
      sheetName: 'Sheet1',
      row: 0,
      column: 0,
      writeValue: true,
      cell: { value: 'second' },
      style: { bold: true },
    })
    expect(edits[1]).toMatchObject({
      row: 1,
      column: 1,
      cell: { value: '', formula: '=SUM(A1:A1)' },
    })
  })

  it('dryRun setCells validates everything but journals nothing', async () => {
    const { XlsxSession } = await import('../src/xlsx/session.js')
    const session = await XlsxSession.open(join(root, 'book.xlsx'), root, makeStubIo())
    const { journaled } = session.setCells(
      { sheet: 'Sheet1', cells: [{ ref: 'A1', value: 1 }] },
      true,
    )
    expect(journaled).toBe(1)
    expect(session.meta().dirty).toBe(false)
    expect(() =>
      session.setCells(
        {
          sheet: 'Sheet1',
          cells: [
            { ref: 'A1', value: 1 },
            { ref: 'ZZZZ99', value: 2 },
          ],
        },
        true,
      ),
    ).toThrow(/Invalid A1-style range/)
  })

  it('rich runs join into the cell value and ride the journal', async () => {
    const { XlsxSession } = await import('../src/xlsx/session.js')
    const session = await XlsxSession.open(join(root, 'book.xlsx'), root, makeStubIo())
    session.setCells({
      sheet: 'Sheet1',
      cells: [
        {
          ref: 'A1',
          rich: [
            { text: 'Big ', bold: true, italic: false, underline: false, strikethrough: false },
            { text: 'news', bold: false, italic: true, underline: false, strikethrough: false },
          ],
        },
      ],
    })
    await session.save()
    const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
    expect(edits[0]).toMatchObject({
      writeValue: true,
      cell: { value: 'Big news' },
      rich: [
        { text: 'Big ', bold: true },
        { text: 'news', italic: true },
      ],
    })
  })

  it('re-reading after save reflects the saved workbook bytes', async () => {
    const { access } = await import('node:fs/promises')
    const { XlsxSession } = await import('../src/xlsx/session.js')
    const session = await XlsxSession.open(join(root, 'book.xlsx'), root, makeStubIo())
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'written' }] })
    const result = await session.save(join(root, 'written.xlsx'))
    expect(result.path).toBe(join(root, 'written.xlsx'))
    expect(result.unchanged).toBe(false)
    await expect(access(join(root, 'written.xlsx'))).resolves.toBeUndefined()
    expect(Buffer.isBuffer(await readFile(join(root, 'written.xlsx')))).toBe(true)
  })
})

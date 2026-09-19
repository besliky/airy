// MCP-level coverage for the read-only .pdf session (PAR-002): open_document
// routes .pdf through pdfjs text extraction (@airy-office/file-parse), the
// read shows the text, and saving is refused. The fixture is a minimal
// single-page PDF written by hand (same shape as file-parse's own tests).
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { buildServer } from '../src/index.js'
import { WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'

interface CallResult {
  isError?: boolean
  content?: Array<{ type: string; text?: string }>
  structuredContent?: Record<string, unknown>
}

interface Session {
  client: Client
  close: () => Promise<void>
}

/** minimal but valid one-page PDF with one text line (pdfjs parses it) */
function buildPdfFixture(text: string): Uint8Array {
  const stream = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`
  const bodies = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let out = '%PDF-1.4\n'
  const offsets: number[] = [0]
  for (let i = 0; i < bodies.length; i++) {
    offsets.push(out.length)
    out += `${i + 1} 0 obj\n${bodies[i]}\nendobj\n`
  }
  const xrefStart = out.length
  out += `xref\n0 ${bodies.length + 1}\n0000000000 65535 f \n`
  for (let i = 1; i <= bodies.length; i++) {
    out += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`
  }
  out += `trailer\n<< /Size ${bodies.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`
  return new TextEncoder().encode(out)
}

let root: string
let previousRoot: string | undefined

async function connectSession(): Promise<Session> {
  const server = buildServer()
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'pdf-test-client', version: '0.0.1' })
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
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-pdf-'))
  process.env[WORKSPACE_ROOT_ENV] = root
  await writeFile(join(root, 'report.pdf'), buildPdfFixture('Hello MCP pdf parsing'))
})

afterAll(async () => {
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

describe('pdf read-only session over MCP', () => {
  it('opens .pdf read-only, extracts text and refuses to save', async () => {
    const { client, close } = await connectSession()
    try {
      const opened = await call(client, 'open_document', { path: 'report.pdf' })
      expect(opened.isError).toBeFalsy()
      expect(opened.structuredContent?.kind).toBe('text')
      expect(opened.structuredContent?.format).toBe('pdf')
      expect(opened.structuredContent?.editable).toBe(false)
      const warnings = (opened.structuredContent?.warnings as string[] | undefined) ?? []
      expect(String(warnings[0] ?? '')).toContain('read-only')
      const handle = String(opened.structuredContent?.handle)

      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      expect(text(read)).toContain('Hello MCP pdf parsing')
      expect(text(read)).toContain('editable: false')
      expect(text(read)).toContain('pages are separated by blank lines')

      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBe(true)
      expect(text(saved)).toContain('read-only text session (.pdf)')

      const insert = await call(client, 'insert_content', {
        handle,
        text: 'nope',
        slide: 0,
      })
      expect(insert.isError).toBe(true)
      await call(client, 'close_document', { handle })
    } finally {
      await close()
    }
  })

  it('fails with an actionable error on a corrupt pdf', async () => {
    const { client, close } = await connectSession()
    try {
      await writeFile(join(root, 'broken.pdf'), Buffer.from('%PDF-1.4 garbage'))
      const opened = await call(client, 'open_document', { path: 'broken.pdf' })
      expect(opened.isError).toBe(true)
      expect(text(opened)).toContain('Cannot extract text')
    } finally {
      await close()
    }
  })
})

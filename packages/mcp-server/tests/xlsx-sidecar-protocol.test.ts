// Protocol codec + process tests for the headless sidecar client. The codec
// functions are pure; the spawn/NDJSON behavior runs against a fake sidecar
// implemented as an executable Node script (the client spawns its binary path
// directly), so request/response correlation, stdout-noise tolerance and
// error propagation are covered without building the Rust binary.
import { chmod, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { encodeRequest, parseSidecarLine, XlsxSidecarClient } from '../src/xlsx/sidecar-client.js'

describe('sidecar protocol codec', () => {
  it('encodes version, requestId and command on one NDJSON line', () => {
    const line = encodeRequest('id-1', { command: 'open', path: '/tmp/book.xlsx' })
    expect(line.endsWith('\n')).toBe(true)
    expect(line.indexOf('\n')).toBe(line.length - 1) // exactly one line
    const parsed = JSON.parse(line) as Record<string, unknown>
    expect(parsed.version).toBe(1)
    expect(parsed.requestId).toBe('id-1')
    expect(parsed.command).toBe('open')
    expect(parsed.path).toBe('/tmp/book.xlsx')
  })

  it('parses ok and error responses', () => {
    const ok = parseSidecarLine('{"version":1,"requestId":"r1","ok":true,"result":{"closed":true}}')
    expect(ok).toMatchObject({ version: 1, requestId: 'r1', ok: true })
    const err = parseSidecarLine(
      '{"version":1,"requestId":"r2","ok":false,"error":{"code":"invalid_request","message":"unknown session"}}',
    )
    expect(err?.ok).toBe(false)
    expect(err?.error?.code).toBe('invalid_request')
  })

  it('treats non-JSON, wrong-version and malformed lines as noise', () => {
    expect(parseSidecarLine('Unexpected type (empty) in shared strings')).toBeNull()
    expect(parseSidecarLine('{"version":2,"requestId":"r","ok":true}')).toBeNull()
    expect(parseSidecarLine('{"version":1,"requestId":7,"ok":true}')).toBeNull()
    expect(parseSidecarLine('{"version":1,"requestId":"r"}')).toBeNull()
    expect(parseSidecarLine('[]')).toBeNull()
  })
})

// Fake sidecar: replies per the v1 wire protocol, emits one noise line on
// stdout at startup, stays silent for sheetId "slow", and fails reads for
// sessionId "missing".
const FAKE_SIDECAR = `#!/usr/bin/env node
const readline = require('node:readline')
process.stdout.write('Unexpected type (empty) in shared strings\\n')
readline.createInterface({ input: process.stdin }).on('line', (line) => {
  let req
  try { req = JSON.parse(line) } catch { return }
  const { requestId, command, ...rest } = req
  const reply = (payload) =>
    process.stdout.write(JSON.stringify({ version: 1, requestId, ...payload }) + '\\n')
  if (req.version !== 1) {
    return reply({ ok: false, error: { code: 'unsupported_version', message: 'bad version' } })
  }
  if (command === 'open') {
    return reply({
      ok: true,
      result: {
        sessionId: 'fake-1',
        name: 'book',
        entryCount: 1,
        sheets: [{ id: 'sheet-0', name: 'S1', rowCount: 3, columnCount: 2 }],
        activeTab: 0,
      },
    })
  }
  if (command === 'read_range') {
    if (rest.sessionId === 'missing') {
      return reply({ ok: false, error: { code: 'invalid_request', message: 'unknown session' } })
    }
    if (rest.sheetId === 'slow') return // never replies
    return reply({
      ok: true,
      result: { cells: [{ row: 0, column: 0, value: 'x' }], indexingComplete: true },
    })
  }
  if (command === 'close') return reply({ ok: true, result: { closed: true } })
  if (command === 'cancel') return reply({ ok: true, result: { cancelled: true } })
  return reply({ ok: false, error: { code: 'invalid_request', message: 'unknown command' } })
})
`

describe('XlsxSidecarClient against a fake sidecar process', () => {
  let dir: string
  let client: XlsxSidecarClient

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'airy-fake-sidecar-'))
    const script = join(dir, 'fake-sidecar')
    await writeFile(script, FAKE_SIDECAR, 'utf8')
    await chmod(script, 0o755)
    client = new XlsxSidecarClient(script)
  })

  afterAll(async () => {
    client.stop()
    await rm(dir, { recursive: true, force: true })
  })

  it('correlates responses by requestId across a real process', async () => {
    const opened = (await client.open('/tmp/book.xlsx')) as { sessionId: string }
    expect(opened.sessionId).toBe('fake-1')
    const range = (await client.readRange({
      sessionId: opened.sessionId,
      sheetId: 'sheet-0',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    })) as { cells: Array<{ value: string }> }
    expect(range.cells[0]?.value).toBe('x')
    await expect(client.close(opened.sessionId)).resolves.toBeUndefined()
  })

  it('propagates sidecar error messages as rejections', async () => {
    await expect(
      client.readRange({
        sessionId: 'missing',
        sheetId: 'sheet-0',
        range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      }),
    ).rejects.toThrow('unknown session')
  })

  it('fails pending requests when the process is stopped', async () => {
    const pending = client.readRange({
      sessionId: 'fake-1',
      sheetId: 'slow',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    })
    const expectation = expect(pending).rejects.toThrow('XLSX sidecar stopped.')
    client.stop()
    await expectation
  })
})

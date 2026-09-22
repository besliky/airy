// End-to-end MCP integration for the headless html tools (PAR-004): a real
// Client over an InMemoryTransport pair drives the full open -> read ->
// insert -> apply_ops -> save -> reopen cycle, the parse5 structure summary,
// byte-level round-trips (BOM, EOLs, untouched lines), the failure paths
// (missing file, outside the workspace root, undeclared non-UTF-8, binary,
// size caps) and the save fences.
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
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

let root: string
let previousRoot: string | undefined

async function connectSession(): Promise<Session> {
  const server = buildServer()
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  const client = new Client({ name: 'html-test-client', version: '0.0.1' })
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

const LF_FIXTURE = [
  '<!DOCTYPE html>',
  '<html>',
  '<head>',
  '  <meta charset="utf-8">',
  '  <title>Quarterly Report</title>',
  '</head>',
  '<body>',
  '  <h1>Quarterly Report</h1>',
  '  <p>Revenue grew by <strong>12 percent</strong> year over year.</p>',
  '  <h2>Detail</h2>',
  '  <ul><li><a href="detail.html">Full detail</a></li></ul>',
  '  <a href="https://example.com">External source</a>',
  '</body>',
  '</html>',
  '',
].join('\n')

async function openFixture(
  client: Client,
  name = 'page.html',
  bytes: Uint8Array = new TextEncoder().encode(LF_FIXTURE),
): Promise<string> {
  await writeFile(join(root, name), bytes)
  const opened = await call(client, 'open_document', { path: name })
  expect(opened.isError).toBeFalsy()
  return String(opened.structuredContent?.handle)
}

// ISOLATION CONTRACT (TEST-705): this suite repoints the process-wide
// workspace root (`process.env[WORKSPACE_ROOT_ENV]`) for its whole lifetime,
// restoring the previous value in afterAll. That is only safe because vitest
// runs test FILES in isolation — one file per worker by default. Do not run
// this suite with `isolate: false`, `singleFork`, or any pool that shares one
// process across files: another file reading the root concurrently (or the
// mid-suite root-drift cases below) would race the shared env var. The CI
// config keeps the default isolation; no describe.sequential needed because
// each file owns its own temp root.
beforeAll(async () => {
  previousRoot = process.env[WORKSPACE_ROOT_ENV]
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-html-'))
  process.env[WORKSPACE_ROOT_ENV] = root
})

afterAll(async () => {
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

describe('html tools over MCP', () => {
  it('serves html through the shared document tools with a parse5 structure summary', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const opened = await call(client, 'open_document', { path: 'page.html' })
      expect(opened.structuredContent?.kind).toBe('html')
      expect(opened.structuredContent?.headingCount).toBe(2)
      expect(opened.structuredContent?.linkCount).toBe(2)
      expect(opened.structuredContent?.title).toBe('Quarterly Report')

      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      const body = text(read)
      expect(body).toContain('Title: Quarterly Report.')
      // headings and links with line positions (0-based)
      expect(body).toContain('1|7|h1|Quarterly Report')
      expect(body).toContain('2|9|h2|Detail')
      expect(body).toContain(`1|10|Full detail -> detail.html`)
      expect(body).toContain('2|11|External source -> https://example.com')
      expect(body).toContain('Line endings: LF.')
      // blocks/range select lines for html sessions
      const ranged = await call(client, 'read_document', { handle, range: { start: 9, end: 9 } })
      expect(text(ranged)).toContain('Selected 1 line(s)')
      expect(text(ranged)).toContain('<h2>Detail</h2>')
      // the workbook reader rejects html handles with a pointer
      const wrong = await call(client, 'read_workbook', { handle })
      expect(wrong.isError).toBe(true)
      expect(text(wrong)).toContain('read_document')
    } finally {
      await close()
    }
  })

  it('runs the full edit cycle: open, read, insert before </body>, apply_ops, save, reopen', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      // insert a section verbatim before the closing body tag (DOC-1506: the
      // marker names the </body> line, so the fragment must stay INSIDE body)
      const inserted = await call(client, 'insert_content', {
        handle,
        html: '  <h2>Outlook</h2>\n  <p>Next quarter looks promising.</p>',
        marker: '</body>',
      })
      expect(inserted.isError).toBeFalsy()
      expect(inserted.structuredContent?.inserted).toBe(2)
      // landed after line 11, i.e. BEFORE the closing tag on line 12
      expect(inserted.structuredContent?.at).toBe(11)
      expect(text(inserted)).toContain('before line 12')

      // line ops: retarget a link, fix a typo, delete the old list line.
      // the window is 12..14: line 15 is the phantom after the final
      // newline and no longer an addressable index (BUG-1102)
      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [
          { op: 'findReplace', find: 'detail.html', replace: 'full-detail.html' },
          { op: 'findReplace', find: 'promising', replace: 'excellent', from: 12, to: 14 },
          { op: 'deleteLines', from: 11, to: 11 },
        ],
      })
      expect(ops.isError).toBeFalsy()
      expect(String(ops.structuredContent?.summary)).toContain('findReplace: matched 1')

      const saved = await call(client, 'save_document', { handle, path: 'edited.html' })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.unchanged).toBe(false)

      const onDisk = (await readFile(join(root, 'edited.html'))).toString('utf8')
      expect(onDisk).toContain('<h2>Outlook</h2>')
      expect(onDisk).toContain('Next quarter looks excellent.')
      expect(onDisk).toContain('full-detail.html')
      expect(onDisk).not.toContain('External source')
      // the appended section sits inside the body element: before </body>,
      // not between </body> and </html> (DOC-1506)
      const bodyClose = onDisk.indexOf('</body>')
      expect(bodyClose).toBeGreaterThan(onDisk.indexOf('<h2>Outlook</h2>'))
      expect(onDisk.indexOf('</html>')).toBeGreaterThan(bodyClose)
      // the reopened copy re-scans structure from the edited markup
      const reopened = await call(client, 'open_document', { path: 'edited.html' })
      expect(reopened.structuredContent?.headingCount).toBe(3)
      expect(reopened.structuredContent?.linkCount).toBe(1)
    } finally {
      await close()
    }
  })

  it('keeps a marker carrying more than the bare closing tag on the after-the-line side', async () => {
    const { client, close } = await connectSession()
    try {
      // only a line that is JUST a structural closing tag flips to the
      // before-side; a line with other markup keeps the documented behavior
      const handle = await openFixture(
        client,
        'inline-close.html',
        new TextEncoder().encode('<body>\n<p>Last.</p></body>\n</html>\n'),
      )
      const inserted = await call(client, 'insert_content', {
        handle,
        html: '<p>New.</p>',
        marker: '</body>',
      })
      expect(inserted.isError).toBeFalsy()
      // marker line is 1 ("<p>Last.</p></body>"): content goes AFTER it
      expect(inserted.structuredContent?.at).toBe(1)
      expect(text(inserted)).toContain('after marker line 1')
    } finally {
      await close()
    }
  })

  it('round-trips bytes verbatim on a zero-edit save', async () => {
    const { client, close } = await connectSession()
    try {
      const bytes = Buffer.concat([
        Buffer.from([0xef, 0xbb, 0xbf]),
        Buffer.from(
          '<!DOCTYPE html>\r\n<html>\r\n<body><p>Hi.</p>\r\n</body>\r\n</html>\r\n',
          'utf8',
        ),
      ])
      const handle = await openFixture(client, 'crlf-bom.html', bytes)
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.unchanged).toBe(true)
      expect(Buffer.compare(await readFile(join(root, 'crlf-bom.html')), bytes)).toBe(0)
    } finally {
      await close()
    }
  })

  it('keeps untouched lines byte-identical through a targeted edit (BOM, mixed EOLs)', async () => {
    const { client, close } = await connectSession()
    try {
      const original = Buffer.from(
        '<p>Line A</p>\r\n<p>Line B</p>\n<p>Line C</p>\r\n<p>Line D</p>\n',
        'utf8',
      )
      const handle = await openFixture(client, 'mixed.html', original)
      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'replaceLines', from: 2, to: 2, text: '<p>Line C1</p>\n<p>Line C2</p>' }],
      })
      expect(ops.isError).toBeFalsy()
      await call(client, 'save_document', { handle })
      const expected = Buffer.from(
        '<p>Line A</p>\r\n<p>Line B</p>\n<p>Line C1</p>\r\n<p>Line C2</p>\r\n<p>Line D</p>\n',
        'utf8',
      )
      expect(Buffer.compare(await readFile(join(root, 'mixed.html')), expected)).toBe(0)
    } finally {
      await close()
    }
  })

  it('preserves the BOM, the EOL style and verbatim fragments for inserts', async () => {
    const { client, close } = await connectSession()
    try {
      const bytes = Buffer.concat([
        Buffer.from([0xef, 0xbb, 0xbf]),
        Buffer.from('<body>\r\n<p>Body.</p>\r\n</body>', 'utf8'),
      ])
      const handle = await openFixture(client, 'bom.html', bytes)
      // verbatim: the fragment lands exactly as written (modulo EOLs); a
      // custom attribute survives untouched
      const inserted = await call(client, 'insert_content', {
        handle,
        html: '<p data-kept="1">Kept</p>',
        at: -1,
      })
      expect(inserted.isError).toBeFalsy()
      expect(text(inserted)).toContain('verbatim')
      await call(client, 'save_document', { handle })
      const saved = await readFile(join(root, 'bom.html'))
      expect(saved[0]).toBe(0xef)
      expect(saved.subarray(3).toString('utf8')).toBe(
        '<p data-kept="1">Kept</p>\r\n<body>\r\n<p>Body.</p>\r\n</body>',
      )
    } finally {
      await close()
    }
  })

  it('gains a line break on the previous last line when appending to a file without one', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(
        client,
        'notail.html',
        Buffer.from('<p>Alpha</p>\n<p>Beta</p>', 'utf8'),
      )
      const inserted = await call(client, 'insert_content', { handle, html: '<p>Gamma</p>' })
      expect(inserted.isError).toBeFalsy()
      expect(text(inserted)).toContain('gained one')
      await call(client, 'save_document', { handle })
      expect((await readFile(join(root, 'notail.html'))).toString('utf8')).toBe(
        '<p>Alpha</p>\n<p>Beta</p>\n<p>Gamma</p>',
      )
    } finally {
      await close()
    }
  })

  it('opens declared legacy charsets and converts edits to UTF-8 with a warning', async () => {
    const { client, close } = await connectSession()
    try {
      // windows-1252: 0xE9 is e-acute; declared via <meta charset>
      const bytes = Buffer.concat([
        Buffer.from(
          '<html><head><meta charset="windows-1252"></head><body><p>caf\xe9</p></body></html>',
          'latin1',
        ),
      ])
      const handle = await openFixture(client, 'legacy.html', bytes)
      const opened = await call(client, 'open_document', { path: 'legacy.html' })
      expect(String((opened.structuredContent?.warnings as string[] | undefined)?.[0])).toContain(
        'windows-1252',
      )
      // the open summary keeps a space between the warning sentence and the
      // Handle pointer (they used to glue: "…original bytes).Handle: …")
      expect(text(opened)).toContain('original bytes). Handle:')
      const read = await call(client, 'read_document', { handle })
      expect(text(read)).toContain('café')
      await call(client, 'insert_content', { handle, html: '<p>More</p>' })
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      const warnings = (saved.structuredContent?.warnings as string[] | undefined) ?? []
      expect(warnings.join(' ')).toContain('UTF-8')
      expect(warnings.join(' ')).toContain('rewritten')
      // the saved bytes are UTF-8 AND the declaration says so: a browser (or
      // any declaration-trusting decoder) must not read the file as mojibake
      const savedBytes = await readFile(join(root, 'legacy.html'))
      const savedText = savedBytes.toString('utf8')
      expect(savedText).toContain('<meta charset="utf-8">')
      expect(savedText).not.toContain('windows-1252')
      const declared = /<meta[^>]+charset\s*=\s*["']?\s*([a-z0-9_.:-]+)/i.exec(savedText)?.[1]
      expect(declared?.toLowerCase()).toBe('utf-8')
      expect(new TextDecoder(declared ?? 'utf-8', { fatal: true }).decode(savedBytes)).toContain(
        'café',
      )
    } finally {
      await close()
    }
  })

  it('rejects an invalid ops batch atomically with a clear message', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const bad = await call(client, 'apply_ops', {
        handle,
        ops: [
          { op: 'findReplace', find: 'Revenue', replace: 'Profit' },
          { op: 'deleteLines', from: 2, to: 99 },
        ],
      })
      expect(bad.isError).toBe(true)
      expect(text(bad)).toContain('deleteLines')
      expect(text(bad)).toContain('nothing was applied')
      const saved = await call(client, 'save_document', { handle, path: 'atomic.html' })
      expect(saved.structuredContent?.unchanged).toBe(true)
    } finally {
      await close()
    }
  })

  it('rejects unknown op fields, unknown ops and bad insert positions', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const unknownField = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'deleteLines', from: 0, to: 0, bogus: 1 }],
      })
      expect(unknownField.isError).toBe(true)
      expect(text(unknownField)).toContain('unknown field "bogus"')
      const unknownOp = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'setFont', target: { nodeType: 'paragraph' }, italic: true }],
      })
      expect(unknownOp.isError).toBe(true)
      expect(text(unknownOp)).toContain('html sessions accept')
      const noMarker = await call(client, 'insert_content', {
        handle,
        html: '<p>x</p>',
        marker: 'does-not-exist',
      })
      expect(noMarker.isError).toBe(true)
      expect(text(noMarker)).toContain('does not match any line')
      const badAt = await call(client, 'insert_content', { handle, html: '<p>x</p>', at: 99 })
      expect(badAt.isError).toBe(true)
      expect(text(badAt)).toContain('out of range')
      // afterHeading is markdown-only: the html branch must reject it instead
      // of silently ignoring the position
      const mdOnly = await call(client, 'insert_content', {
        handle,
        html: '<p>x</p>',
        afterHeading: 1,
      })
      expect(mdOnly.isError).toBe(true)
      expect(text(mdOnly)).toContain('afterHeading is a markdown-session option')
      // `text` is markdown-only too: the html branch must reject it instead of
      // silently discarding it (the markdown branch rejects `html` the same way)
      const wrongPayload = await call(client, 'insert_content', {
        handle,
        text: '<p>x</p>',
      })
      expect(wrongPayload.isError).toBe(true)
      expect(text(wrongPayload)).toContain('pass the fragment in `html`, not `text`')
    } finally {
      await close()
    }
  })

  it('fails clearly for missing files, outside-root paths, undeclared non-UTF-8 and binary', async () => {
    const { client, close } = await connectSession()
    try {
      const missing = await call(client, 'open_document', { path: 'nope.html' })
      expect(missing.isError).toBe(true)
      expect(text(missing)).toContain('Cannot read')
      const outside = await call(client, 'open_document', { path: '../../etc/passwd' })
      expect(outside.isError).toBe(true)
      expect(text(outside)).toContain('outside the workspace root')
      // invalid UTF-8 without a charset declaration is refused
      await writeFile(
        join(root, 'broken.html'),
        Buffer.from('<html><body>ok \x80 text</body></html>', 'latin1'),
      )
      const broken = await call(client, 'open_document', { path: 'broken.html' })
      expect(broken.isError).toBe(true)
      expect(text(broken)).toContain('not valid UTF-8')
      // NUL bytes mean binary
      await writeFile(join(root, 'binary.html'), Buffer.from('abc\0def', 'latin1'))
      const binary = await call(client, 'open_document', { path: 'binary.html' })
      expect(binary.isError).toBe(true)
      expect(text(binary)).toContain('NUL')
      // unsupported extension still lists what open_document takes
      const ext = await call(client, 'open_document', { path: 'file.css' })
      expect(ext.isError).toBe(true)
      expect(text(ext)).toContain('Unsupported file type')
    } finally {
      await close()
    }
  })

  it('enforces the size limits: 8 MiB open cap, 30k read budget, 100-op batch cap', async () => {
    const { client, close } = await connectSession()
    try {
      await writeFile(join(root, 'huge.html'), Buffer.alloc(8 * 1024 * 1024 + 1, 0x61))
      const huge = await call(client, 'open_document', { path: 'huge.html' })
      expect(huge.isError).toBe(true)
      expect(text(huge)).toContain('8 MiB')

      const handle = await openFixture(
        client,
        'long.html',
        new TextEncoder().encode(`<!DOCTYPE html><body>${'x'.repeat(31_000)}\nend</body>`),
      )
      const read = await call(client, 'read_document', { handle })
      expect(text(read)).toContain('output truncated at 30000 characters')
      const ranged = await call(client, 'read_document', { handle, range: { start: 1, end: 1 } })
      expect(text(ranged)).toContain('end')

      const tooMany = await call(client, 'apply_ops', {
        handle,
        ops: Array.from({ length: 101 }, () => ({ op: 'deleteLines', from: 0, to: 0 })),
      })
      expect(tooMany.isError).toBe(true)
    } finally {
      await close()
    }
  })

  it('caps the line count at open: a file of bare EOLs is refused before the model is built', async () => {
    const { client, close } = await connectSession()
    try {
      // 2,000,001 lines in ~2 MB: under the byte cap, but the line model
      // would be millions of objects. The refusal must come from the line
      // cap (counted before splitLines), not from the byte cap.
      await writeFile(join(root, 'eol-flood.html'), Buffer.from('\n'.repeat(2_000_000)))
      const flooded = await call(client, 'open_document', { path: 'eol-flood.html' })
      expect(flooded.isError).toBe(true)
      expect(text(flooded)).toContain('2000001 lines')
      expect(text(flooded)).toContain('cap at 2000000 lines')
      expect(text(flooded)).not.toContain('8 MiB')
    } finally {
      await close()
    }
  })

  it('rejects an enormous read range fast, before allocating the index array', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      // before the span cap this loop built the whole index array first and
      // hung/OOMed the server; it must now fail fast with the cap message
      const exploded = await call(client, 'read_document', {
        handle,
        range: { start: 0, end: Number.MAX_SAFE_INTEGER },
      })
      expect(exploded.isError).toBe(true)
      expect(text(exploded)).toContain('the cap is 10000 per read')
      // out-of-range counting still works without materializing the tail
      const beyond = await call(client, 'read_document', {
        handle,
        range: { start: 10, end: 20 },
      })
      expect(beyond.isError).toBe(true)
      expect(text(beyond)).toContain('out of range')
    } finally {
      await close()
    }
  })

  it('caps the heading list and counts the structure summary toward the 30k budget', async () => {
    const { client, close } = await connectSession()
    try {
      // 300 headings plus a 31k text line: without the caps the heading list
      // was uncapped (a heading-heavy file could flood the answer) and the
      // text got its own full 30k on top of the structure block
      const flood = [
        '<html><body>',
        ...Array.from({ length: 300 }, (_, i) => `<h2>heading ${String(i)}</h2>`),
        `<p>${'x'.repeat(31_000)}</p>`,
        '</body></html>',
      ].join('\n')
      const handle = await openFixture(client, 'flood.html', new TextEncoder().encode(flood))
      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      const body = text(read)
      expect(body).toContain('(first 200 of 300 - use range reads for the rest)')
      expect(body).toContain('|h2|heading 0')
      expect(body).not.toContain('|h2|heading 250')
      // the structure block counts toward the budget: the 31k text no longer
      // gets its own full 30k slice on top of the summary
      expect(body).toContain('output truncated at 30000 characters')
      expect(body.length).toBeLessThan(31_500)
    } finally {
      await close()
    }
  })

  it('guards saves: refuses existing targets without overwrite, fences external writers', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      await writeFile(join(root, 'existing.html'), 'precious')
      const refused = await call(client, 'save_document', { handle, path: 'existing.html' })
      expect(refused.isError).toBe(true)
      expect(text(refused)).toContain('already exists')
      expect((await readFile(join(root, 'existing.html'))).toString('utf8')).toBe('precious')
      const overwrote = await call(client, 'save_document', {
        handle,
        path: 'existing.html',
        overwrite: true,
      })
      expect(overwrote.isError).toBeFalsy()

      const handle2 = await openFixture(client, 'fenced.html')
      await call(client, 'insert_content', { handle: handle2, html: '<p>edit</p>' })
      await writeFile(join(root, 'fenced.html'), 'externally rewritten')
      const fenced = await call(client, 'save_document', { handle: handle2 })
      expect(fenced.isError).toBe(true)
      expect(text(fenced)).toContain('changed on disk')
    } finally {
      await close()
    }
  })

  it('opens the short .htm extension and closes html sessions', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client, 'page.htm')
      const closed = await call(client, 'close_document', { handle })
      expect(closed.isError).toBeFalsy()
      expect(closed.structuredContent?.kind).toBe('html')
      const after = await call(client, 'read_document', { handle })
      expect(after.isError).toBe(true)
      expect(text(after)).toContain('Unknown document handle')
    } finally {
      await close()
    }
  })

  it('rejects findReplace text containing line breaks (line-model invariant)', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const multilineFind = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'Revenue\nDetail', replace: 'x' }],
      })
      expect(multilineFind.isError).toBe(true)
      expect(text(multilineFind)).toContain('find must not contain line breaks')
      // a replace with an embedded EOL would leave a line break inside one
      // line object; it must refuse and point at the multi-line ops instead
      const multilineReplace = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'Revenue', replace: 'Sales\nGrowth' }],
      })
      expect(multilineReplace.isError).toBe(true)
      expect(text(multilineReplace)).toContain('replace must not contain line breaks')
      expect(text(multilineReplace)).toContain('insertLines or replaceLines')
      // single-line replacements keep working
      const ok = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'Revenue', replace: 'Sales' }],
      })
      expect(ok.isError).toBeFalsy()
    } finally {
      await close()
    }
  })

  it('supports dryRun ops and case-insensitive find/replace', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const dry = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'revenue', replace: 'sales', matchCase: false }],
        dryRun: true,
      })
      expect(dry.isError).toBeFalsy()
      expect(text(dry)).toContain('dry run, nothing applied')
      const applied = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'revenue', replace: 'sales', matchCase: false }],
      })
      expect(applied.isError).toBeFalsy()
      const read = await call(client, 'read_document', { handle })
      expect(text(read)).toContain('sales grew by')
    } finally {
      await close()
    }
  })

  it('case-insensitive findReplace survives the length-changing İ fold (BUG-1101)', async () => {
    const { client, close } = await connectSession()
    try {
      // audit repro on the html surface: İ (U+0130) folds to two code
      // units, which used to shift lowered indices and corrupt the tail
      const handle = await openFixture(
        client,
        'turkish.html',
        Buffer.from('<p>İstanbul kelime not</p>\n', 'utf8'),
      )
      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'kelime', replace: 'WORD', matchCase: false }],
      })
      expect(ops.isError).toBeFalsy()
      expect(String(ops.structuredContent?.summary)).toContain('findReplace: matched 1')
      await call(client, 'save_document', { handle })
      expect((await readFile(join(root, 'turkish.html'))).toString('utf8')).toBe(
        '<p>İstanbul WORD not</p>\n',
      )
    } finally {
      await close()
    }
  })

  it('default-append keeps the trailing-newline shape and skips no leading newline (BUG-1102)', async () => {
    const { client, close } = await connectSession()
    try {
      // file WITH a trailing newline: the fragment lands after the last
      // real line, the final \n survives, no blank line appears
      const handle = await openFixture(
        client,
        'tail.html',
        Buffer.from('<p>a</p>\n<p>b</p>\n', 'utf8'),
      )
      const opened = await call(client, 'open_document', { path: 'tail.html' })
      expect(opened.structuredContent?.lineCount).toBe(2)
      const atEnd = await call(client, 'insert_content', { handle, html: '<p>c</p>' })
      expect(atEnd.isError).toBeFalsy()
      expect(atEnd.structuredContent?.lineCount).toBe(3)
      await call(client, 'save_document', { handle })
      expect(
        Buffer.compare(
          await readFile(join(root, 'tail.html')),
          Buffer.from('<p>a</p>\n<p>b</p>\n<p>c</p>\n'),
        ),
      ).toBe(0)

      // empty file: no leading newline (default and at:-1 agree)
      const handle2 = await openFixture(client, 'empty.html', Buffer.from('', 'utf8'))
      await call(client, 'insert_content', { handle: handle2, html: '<p>hello</p>' })
      await call(client, 'save_document', { handle: handle2 })
      expect((await readFile(join(root, 'empty.html'))).toString('utf8')).toBe('<p>hello</p>\n')
    } finally {
      await close()
    }
  })

  it('confines saves to the workspace root captured at open, not a later one', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client, 'pinned.html')
      await call(client, 'insert_content', { handle, html: '<p>edit</p>' })
      // drift AIRY_WORKSPACE_ROOT after open: before the fix the default
      // save re-confined the absolute opened path against the NEW root and
      // failed with PathOutsideWorkspaceError; a relative save-as resolved
      // into the wrong directory
      const driftRoot = await mkdtemp(join(tmpdir(), 'airy-mcp-html-drift-'))
      process.env[WORKSPACE_ROOT_ENV] = driftRoot
      try {
        const inPlace = await call(client, 'save_document', { handle })
        expect(inPlace.isError).toBeFalsy()
        expect(String(inPlace.structuredContent?.path)).toBe(join(root, 'pinned.html'))
        const saveAs = await call(client, 'save_document', { handle, path: 'pinned-out.html' })
        expect(saveAs.isError).toBeFalsy()
        // the relative target resolved against the OPEN-time root
        expect(existsSync(join(root, 'pinned-out.html'))).toBe(true)
        expect(existsSync(join(driftRoot, 'pinned-out.html'))).toBe(false)
        expect(
          (await readFile(join(root, 'pinned-out.html'), 'utf8')).includes('<p>edit</p>'),
        ).toBe(true)
      } finally {
        process.env[WORKSPACE_ROOT_ENV] = root
        await rm(driftRoot, { recursive: true, force: true })
      }
    } finally {
      await close()
    }
  })
})

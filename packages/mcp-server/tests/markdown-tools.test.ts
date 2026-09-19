// End-to-end MCP integration for the headless markdown tools (PAR-003):
// a real Client over an InMemoryTransport pair drives the full open -> read
// -> insert -> apply_ops -> save -> reopen cycle, byte-level round-trips
// (BOM, EOLs, untouched lines), the failure paths (missing file, outside the
// workspace root, invalid UTF-8, binary, size caps) and the save fences.
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
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
  const client = new Client({ name: 'markdown-test-client', version: '0.0.1' })
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

/** the shared LF fixture: three headings, one inside a code fence that must not count */
const LF_FIXTURE = [
  '# Quarterly Report',
  '',
  'Revenue grew by 12 percent year over year.',
  'Costs fell slightly.',
  '## Revenue Detail',
  '- EMEA: strong',
  '- APAC: recovering',
  '```js',
  '# not a heading',
  '```',
  '## Outlook',
  'Next quarter looks promising.',
  '',
].join('\n')

async function openFixture(
  client: Client,
  name = 'notes.md',
  bytes: Uint8Array = new TextEncoder().encode(LF_FIXTURE),
): Promise<string> {
  await writeFile(join(root, name), bytes)
  const opened = await call(client, 'open_document', { path: name })
  expect(opened.isError).toBeFalsy()
  return String(opened.structuredContent?.handle)
}

beforeAll(async () => {
  previousRoot = process.env[WORKSPACE_ROOT_ENV]
  root = await mkdtemp(join(tmpdir(), 'airy-mcp-md-'))
  process.env[WORKSPACE_ROOT_ENV] = root
})

afterAll(async () => {
  if (previousRoot === undefined) delete process.env[WORKSPACE_ROOT_ENV]
  else process.env[WORKSPACE_ROOT_ENV] = previousRoot
  await rm(root, { recursive: true, force: true })
})

describe('markdown tools over MCP', () => {
  it('serves markdown through the shared document tools', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      const body = text(read)
      // structure: headings with positions, fence content skipped
      expect(body).toContain('1|0|#|Quarterly Report')
      expect(body).toContain('2|4|##|Revenue Detail')
      expect(body).toContain('3|10|##|Outlook')
      expect(body).not.toContain('|#|not a heading')
      expect(body).toContain('Line endings: LF.')
      expect(body).toContain('Full text (EOLs normalized to LF):')
      // blocks/range select lines for markdown sessions
      const ranged = await call(client, 'read_document', { handle, range: { start: 4, end: 5 } })
      expect(text(ranged)).toContain('Selected 2 line(s)')
      expect(text(ranged)).toContain('## Revenue Detail')
      // the workbook reader rejects markdown handles with a pointer
      const wrong = await call(client, 'read_workbook', { handle })
      expect(wrong.isError).toBe(true)
      expect(text(wrong)).toContain('read_document')
    } finally {
      await close()
    }
  })

  it('runs the full edit cycle: open, read, insert, apply_ops, save, reopen', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const opened = await call(client, 'open_document', { path: 'notes.md' })
      expect(opened.structuredContent?.kind).toBe('markdown')
      // 12 real lines + a trailing newline: the phantom position after the
      // final \n is not a line (BUG-1102)
      expect(opened.structuredContent?.lineCount).toBe(12)
      expect(opened.structuredContent?.headingCount).toBe(3)
      expect(opened.structuredContent?.eol).toBe('\n')

      // insert after heading 2 ("Revenue Detail", line 4)
      const byHeading = await call(client, 'insert_content', {
        handle,
        text: '### EMEA detail\nEverything grew.',
        afterHeading: 2,
      })
      expect(byHeading.isError).toBeFalsy()
      expect(byHeading.structuredContent?.inserted).toBe(2)

      // insert after a marker line
      const byMarker = await call(client, 'insert_content', {
        handle,
        text: '- AMERICAS: record',
        marker: 'APAC: recovering',
      })
      expect(byMarker.isError).toBeFalsy()
      expect(byMarker.structuredContent?.at).toBe(8) // marker shifted to line 8

      // default insert appends at the end (12 real lines + 2 + 1 + 2)
      const atEnd = await call(client, 'insert_content', {
        handle,
        text: 'Signed,\nThe Team',
      })
      expect(atEnd.isError).toBeFalsy()
      expect(atEnd.structuredContent?.lineCount).toBe(17)

      // line ops: rename a bullet, delete the costs line, scoped find/replace
      // (the to bound shrinks by one after the delete: the phantom after the
      // final newline is no longer an addressable index — BUG-1102)
      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [
          { op: 'findReplace', find: 'EMEA: strong', replace: 'EMEA: very strong' },
          { op: 'deleteLines', from: 3, to: 3 },
          { op: 'findReplace', find: 'promising', replace: 'excellent', from: 10, to: 15 },
        ],
      })
      expect(ops.isError).toBeFalsy()
      expect(String(ops.structuredContent?.summary)).toContain('findReplace: matched 1')

      const saved = await call(client, 'save_document', { handle, path: 'edited.md' })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.unchanged).toBe(false)

      const onDisk = await readFile(join(root, 'edited.md'), 'utf8')
      expect(onDisk).toContain('- EMEA: very strong')
      expect(onDisk).toContain('- AMERICAS: record')
      expect(onDisk).not.toContain('Costs fell slightly.')
      expect(onDisk).toContain('Next quarter looks excellent.')
      // reopening the saved copy reflects the new structure
      const reopened = await call(client, 'open_document', { path: 'edited.md' })
      expect(reopened.structuredContent?.headingCount).toBe(4)
    } finally {
      await close()
    }
  })

  it('round-trips bytes verbatim on a zero-edit save', async () => {
    const { client, close } = await connectSession()
    try {
      const bytes = Buffer.concat([
        Buffer.from([0xef, 0xbb, 0xbf]),
        Buffer.from('# Title\r\n\r\nBody line one.\r\nBody line two.\r\nTail line.', 'utf8'),
      ])
      const handle = await openFixture(client, 'crlf-bom.md', bytes)
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(saved.structuredContent?.unchanged).toBe(true)
      expect(Buffer.compare(await readFile(join(root, 'crlf-bom.md')), bytes)).toBe(0)
    } finally {
      await close()
    }
  })

  it('keeps untouched lines byte-identical through a targeted edit (BOM, mixed EOLs)', async () => {
    const { client, close } = await connectSession()
    try {
      // mixed terminators: A/B/D end CRLF, C ends LF, plus a final trailing LF
      const original = Buffer.from('Line A\r\nLine B\nLine C\r\nLine D\nLine E\r\n', 'utf8')
      const handle = await openFixture(client, 'mixed.md', original)
      // replace line 2 ("Line C") with two lines; dominant EOL is CRLF
      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'replaceLines', from: 2, to: 2, text: 'Line C1\nLine C2' }],
      })
      expect(ops.isError).toBeFalsy()
      await call(client, 'save_document', { handle })
      const expected = Buffer.from(
        'Line A\r\nLine B\nLine C1\r\nLine C2\r\nLine D\nLine E\r\n',
        'utf8',
      )
      expect(Buffer.compare(await readFile(join(root, 'mixed.md')), expected)).toBe(0)
    } finally {
      await close()
    }
  })

  it('preserves the BOM and the file EOL style for inserts', async () => {
    const { client, close } = await connectSession()
    try {
      const bytes = Buffer.concat([
        Buffer.from([0xef, 0xbb, 0xbf]),
        Buffer.from('# Title\r\n\r\nBody.\r\n', 'utf8'),
      ])
      const handle = await openFixture(client, 'bom.md', bytes)
      const inserted = await call(client, 'insert_content', {
        handle,
        text: '## New section\nMore text.',
        at: -1,
      })
      expect(inserted.isError).toBeFalsy()
      await call(client, 'save_document', { handle })
      const saved = await readFile(join(root, 'bom.md'))
      // BOM survived and the inserted lines use CRLF
      expect(saved[0]).toBe(0xef)
      expect(saved[1]).toBe(0xbb)
      expect(saved[2]).toBe(0xbf)
      expect(saved.subarray(3).toString('utf8')).toBe(
        '## New section\r\nMore text.\r\n# Title\r\n\r\nBody.\r\n',
      )
    } finally {
      await close()
    }
  })

  it('gains a line break on the previous last line when appending to a file without one', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client, 'notail.md', Buffer.from('Alpha\nBeta', 'utf8'))
      const inserted = await call(client, 'insert_content', { handle, text: 'Gamma' })
      expect(inserted.isError).toBeFalsy()
      expect(text(inserted)).toContain('gained one')
      await call(client, 'save_document', { handle })
      expect((await readFile(join(root, 'notail.md'))).toString('utf8')).toBe('Alpha\nBeta\nGamma')
    } finally {
      await close()
    }
  })

  it('opens UTF-16 BOM files and converts edits to UTF-8 with a warning', async () => {
    const { client, close } = await connectSession()
    try {
      const utf16 = Buffer.from('\ufeff# Wide\n\nBody.', 'utf16le')
      const handle = await openFixture(client, 'wide.md', utf16)
      const opened = await call(client, 'open_document', { path: 'wide.md' })
      expect(opened.structuredContent?.bom).toBe(true)
      expect(String((opened.structuredContent?.warnings as string[] | undefined)?.[0])).toContain(
        'UTF-16',
      )
      // the open summary keeps a space between the warning sentence and the
      // Handle pointer (they used to glue: "…writes UTF-8.Handle: …")
      expect(text(opened)).toContain('writes UTF-8. Handle:')
      await call(client, 'insert_content', { handle, text: 'Tail' })
      const saved = await call(client, 'save_document', { handle })
      expect(saved.isError).toBeFalsy()
      expect(String((saved.structuredContent?.warnings as string[] | undefined)?.[0])).toContain(
        'UTF-8',
      )
      const onDisk = await readFile(join(root, 'wide.md'))
      expect(onDisk[0]).toBe(0xef) // saved as UTF-8 with the BOM re-applied
      expect(onDisk.toString('utf8').endsWith('Tail')).toBe(true)
    } finally {
      await close()
    }
  })

  it('skips YAML front matter when scanning headings', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(
        client,
        'front.md',
        new TextEncoder().encode('---\ntitle: Not a heading\n# also not\n---\n\n# Real\n'),
      )
      const read = await call(client, 'read_document', { handle })
      const body = text(read)
      expect(body).toContain('1|5|#|Real')
      expect(body).not.toContain('|#|also not')
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
      // the valid first op must not have landed: a save is byte-identical
      const saved = await call(client, 'save_document', { handle, path: 'atomic.md' })
      expect(saved.structuredContent?.unchanged).toBe(true)
    } finally {
      await close()
    }
  })

  it('rejects unknown op fields and unknown ops for markdown sessions', async () => {
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
      expect(text(unknownOp)).toContain('markdown sessions accept')
    } finally {
      await close()
    }
  })

  it('reports clear errors for bad insert positions', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client)
      const noMarker = await call(client, 'insert_content', {
        handle,
        text: 'x',
        marker: 'does-not-exist',
      })
      expect(noMarker.isError).toBe(true)
      expect(text(noMarker)).toContain('does not match any line')
      const noHeading = await call(client, 'insert_content', {
        handle,
        text: 'x',
        afterHeading: 4,
      })
      expect(noHeading.isError).toBe(true)
      expect(text(noHeading)).toContain('afterHeading 4 is out of range')
      const badAt = await call(client, 'insert_content', { handle, text: 'x', at: 99 })
      expect(badAt.isError).toBe(true)
      expect(text(badAt)).toContain('out of range')
      const wrongParam = await call(client, 'insert_content', {
        handle,
        html: '<p>docx only</p>',
      })
      expect(wrongParam.isError).toBe(true)
      expect(text(wrongParam)).toContain('pass markdown source in `text`')
    } finally {
      await close()
    }
  })

  it('fails clearly for missing files, outside-root paths, invalid UTF-8 and binary', async () => {
    const { client, close } = await connectSession()
    try {
      const missing = await call(client, 'open_document', { path: 'nope.md' })
      expect(missing.isError).toBe(true)
      expect(text(missing)).toContain('Cannot read')
      const outside = await call(client, 'open_document', { path: '../../etc/passwd' })
      expect(outside.isError).toBe(true)
      expect(text(outside)).toContain('outside the workspace root')
      // a lone UTF-8 continuation byte cannot decode
      await writeFile(join(root, 'broken.md'), Buffer.from([0x61, 0x80, 0x62]))
      const broken = await call(client, 'open_document', { path: 'broken.md' })
      expect(broken.isError).toBe(true)
      expect(text(broken)).toContain('not valid UTF-8')
      // NUL bytes mean binary
      await writeFile(join(root, 'binary.md'), Buffer.from('abc\0def', 'latin1'))
      const binary = await call(client, 'open_document', { path: 'binary.md' })
      expect(binary.isError).toBe(true)
      expect(text(binary)).toContain('NUL')
      // unsupported extension still lists what open_document takes
      const ext = await call(client, 'open_document', { path: 'file.txt' })
      expect(ext.isError).toBe(true)
      expect(text(ext)).toContain('Unsupported file type')
    } finally {
      await close()
    }
  })

  it('enforces the size limits: 8 MiB open cap, 30k read budget, 100-op batch cap', async () => {
    const { client, close } = await connectSession()
    try {
      await writeFile(
        join(root, 'huge.md'),
        Buffer.concat([Buffer.alloc(8 * 1024 * 1024 + 1, 0x61)]),
      )
      const huge = await call(client, 'open_document', { path: 'huge.md' })
      expect(huge.isError).toBe(true)
      expect(text(huge)).toContain('8 MiB')

      const handle = await openFixture(
        client,
        'long.md',
        new TextEncoder().encode(`${'x'.repeat(31_000)}\nend`),
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
      await writeFile(join(root, 'eol-flood.md'), Buffer.from('\n'.repeat(2_000_000)))
      const flooded = await call(client, 'open_document', { path: 'eol-flood.md' })
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
      // a span just over the cap is refused too; at or under it still works
      const over = await call(client, 'read_document', {
        handle,
        range: { start: 0, end: 10_000 },
      })
      expect(over.isError).toBe(true)
      expect(text(over)).toContain('the cap is 10000 per read')
      // out-of-range counting still works without materializing the tail
      const beyond = await call(client, 'read_document', {
        handle,
        range: { start: 10, end: 20 },
      })
      expect(beyond.isError).toBe(true)
      expect(text(beyond)).toContain('out of range')
      const ok = await call(client, 'read_document', { handle, range: { start: 0, end: 11 } })
      expect(text(ok)).toContain('Selected 12 line(s)')
    } finally {
      await close()
    }
  })

  it('caps the heading list and the whole read stays inside the 30k budget (heading flood)', async () => {
    const { client, close } = await connectSession()
    try {
      // one megabyte-long heading line + 250 short ones: without the caps the
      // heading block alone carried the whole file (~100 KB) into the answer
      const flood = [
        `# ${'h'.repeat(100_000)}`,
        ...Array.from({ length: 250 }, (_, i) => `# heading ${String(i)}`),
        '',
      ].join('\n')
      const handle = await openFixture(client, 'flood.md', new TextEncoder().encode(flood))
      const read = await call(client, 'read_document', { handle })
      expect(read.isError).toBeFalsy()
      const body = text(read)
      expect(body).toContain('(first 200 of 251 - use range reads for the rest)')
      expect(body).toContain('2|1|#|heading 0')
      expect(body).toContain('200|199|#|heading 198')
      expect(body).not.toContain('201|')
      // the megabyte heading text is clipped in the list ... and the whole
      // assembly is truncated at the 30k budget
      expect(body).toMatch(/^1\|0\|#\|h{1,80}\.\.\.$/m)
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
      await writeFile(join(root, 'existing.md'), 'precious')
      const refused = await call(client, 'save_document', { handle, path: 'existing.md' })
      expect(refused.isError).toBe(true)
      expect(text(refused)).toContain('already exists')
      expect((await readFile(join(root, 'existing.md'))).toString('utf8')).toBe('precious')
      const overwrote = await call(client, 'save_document', {
        handle,
        path: 'existing.md',
        overwrite: true,
      })
      expect(overwrote.isError).toBeFalsy()

      // fencing: an external writer invalidates an in-place save
      const handle2 = await openFixture(client, 'fenced.md')
      await call(client, 'insert_content', { handle: handle2, text: 'edit' })
      await writeFile(join(root, 'fenced.md'), 'externally rewritten')
      const fenced = await call(client, 'save_document', { handle: handle2 })
      expect(fenced.isError).toBe(true)
      expect(text(fenced)).toContain('changed on disk')
    } finally {
      await close()
    }
  })

  it('opens the long .markdown extension and closes markdown sessions', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client, 'notes.markdown')
      const closed = await call(client, 'close_document', { handle })
      expect(closed.isError).toBeFalsy()
      expect(closed.structuredContent?.kind).toBe('markdown')
      const after = await call(client, 'read_document', { handle })
      expect(after.isError).toBe(true)
      expect(text(after)).toContain('Unknown document handle')
    } finally {
      await close()
    }
  })

  it('labels sessions and builds temp names from the file basename, not the whole path', async () => {
    const { client, close } = await connectSession()
    try {
      // basename (node:path) rather than a '/'-split: on Windows a
      // `\`-separated path never splits, which used to leak the whole path
      // into meta.fileName and into the atomic-save temp file name
      await mkdir(join(root, 'nested/deep'), { recursive: true })
      await writeFile(join(root, 'nested/deep/notes.md'), LF_FIXTURE, 'utf8')
      const opened = await call(client, 'open_document', { path: 'nested/deep/notes.md' })
      expect(opened.isError).toBeFalsy()
      expect(opened.structuredContent?.fileName).toBe('notes.md')
      expect(text(opened)).toContain('Opened notes.md as an editable markdown session')
      // a save into the nested dir exercises the same basename-derived temp
      // name inside the target directory
      const handle = String(opened.structuredContent?.handle)
      const saved = await call(client, 'save_document', { handle, path: 'nested/deep/copy.md' })
      expect(saved.isError).toBeFalsy()
      const onDisk = await readFile(join(root, 'nested/deep/copy.md'), 'utf8')
      expect(onDisk).toContain('# Quarterly Report')
      // no leftover temp files in the target directory
      const entries = await readdir(join(root, 'nested/deep'))
      expect(entries.sort()).toEqual(['copy.md', 'notes.md'])
    } finally {
      await close()
    }
  })

  it('confines saves to the workspace root captured at open, not a later one', async () => {
    const { client, close } = await connectSession()
    try {
      const handle = await openFixture(client, 'pinned.md')
      await call(client, 'insert_content', { handle, text: 'Edit.' })
      // drift AIRY_WORKSPACE_ROOT after open: before the fix the default
      // save re-confined the absolute opened path against the NEW root and
      // failed with PathOutsideWorkspaceError; a relative save-as resolved
      // into the wrong directory
      const driftRoot = await mkdtemp(join(tmpdir(), 'airy-mcp-md-drift-'))
      process.env[WORKSPACE_ROOT_ENV] = driftRoot
      try {
        const inPlace = await call(client, 'save_document', { handle })
        expect(inPlace.isError).toBeFalsy()
        expect(String(inPlace.structuredContent?.path)).toBe(join(root, 'pinned.md'))
        const saveAs = await call(client, 'save_document', { handle, path: 'pinned-out.md' })
        expect(saveAs.isError).toBeFalsy()
        // the relative target resolved against the OPEN-time root
        expect(existsSync(join(root, 'pinned-out.md'))).toBe(true)
        expect(existsSync(join(driftRoot, 'pinned-out.md'))).toBe(false)
        // the appended block landed before the trailing phantom: the file
        // keeps its final newline instead of losing it (BUG-1102)
        expect((await readFile(join(root, 'pinned-out.md'), 'utf8')).endsWith('Edit.\n')).toBe(true)
      } finally {
        process.env[WORKSPACE_ROOT_ENV] = root
        await rm(driftRoot, { recursive: true, force: true })
      }
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
      expect(String(dry.structuredContent?.summary)).toContain('matched 2')
      expect(text(dry)).toContain('dry run, nothing applied')
      const applied = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'revenue', replace: 'sales', matchCase: false }],
      })
      expect(applied.isError).toBeFalsy()
      const read = await call(client, 'read_document', { handle })
      expect(text(read)).toContain('sales grew by 12 percent')
    } finally {
      await close()
    }
  })

  it('case-insensitive findReplace survives the length-changing İ fold (BUG-1101)', async () => {
    const { client, close } = await connectSession()
    try {
      // audit live repro: toLowerCase maps İ (U+0130) to "i" + U+0307, so
      // lowered indices after it shift by one and the replacement used to
      // land mid-word ("İstanbul kWORDnot") with the tail eaten
      const handle = await openFixture(
        client,
        'turkish.md',
        Buffer.from('İstanbul kelime not\n', 'utf8'),
      )
      const ops = await call(client, 'apply_ops', {
        handle,
        ops: [{ op: 'findReplace', find: 'kelime', replace: 'WORD', matchCase: false }],
      })
      expect(ops.isError).toBeFalsy()
      expect(String(ops.structuredContent?.summary)).toContain('findReplace: matched 1')
      await call(client, 'save_document', { handle })
      expect((await readFile(join(root, 'turkish.md'))).toString('utf8')).toBe(
        'İstanbul WORD not\n',
      )
    } finally {
      await close()
    }
  })

  it('default-append into a file with a trailing newline keeps the shape (BUG-1102)', async () => {
    const { client, close } = await connectSession()
    try {
      // audit live repro: "a\nb\n" + insert "c" used to save "a\nb\n\nc" —
      // the phantom after the final \n became a real blank line and the
      // file lost its trailing newline; lineCount reported 3 for 2 lines
      const handle = await openFixture(client, 'tail.md', Buffer.from('a\nb\n', 'utf8'))
      const opened = await call(client, 'open_document', { path: 'tail.md' })
      expect(opened.structuredContent?.lineCount).toBe(2)
      const atEnd = await call(client, 'insert_content', { handle, text: 'c' })
      expect(atEnd.isError).toBeFalsy()
      expect(atEnd.structuredContent?.lineCount).toBe(3)
      expect(text(atEnd)).not.toContain('gained one')
      await call(client, 'save_document', { handle })
      expect(Buffer.compare(await readFile(join(root, 'tail.md')), Buffer.from('a\nb\nc\n'))).toBe(
        0,
      )

      // the op path lands the same way: insertLines after the LAST line
      // (index count-1) inserts before the phantom, not after it
      const handle2 = await openFixture(client, 'tail2.md', Buffer.from('a\r\nb\r\n', 'utf8'))
      const ops = await call(client, 'apply_ops', {
        handle: handle2,
        ops: [{ op: 'insertLines', after: 1, text: 'c' }],
      })
      expect(ops.isError).toBeFalsy()
      await call(client, 'save_document', { handle: handle2 })
      expect((await readFile(join(root, 'tail2.md'))).toString('utf8')).toBe('a\r\nb\r\nc\r\n')

      // the phantom position is not addressable any more (3 real lines, so
      // index 3 — the slot after the final \n — is out of range)
      const phantom = await call(client, 'apply_ops', {
        handle: handle2,
        ops: [{ op: 'deleteLines', from: 3, to: 3 }],
      })
      expect(phantom.isError).toBe(true)
      expect(text(phantom)).toContain('0 <= from <= to < 3')
    } finally {
      await close()
    }
  })

  it('default-append into an empty file adds no leading newline (BUG-1102)', async () => {
    const { client, close } = await connectSession()
    try {
      // the empty file's lone empty line is the whole file: the block lands
      // at the start and the lone line becomes the trailing phantom. The
      // audit repro had a leading "\n"; default and at:-1 must agree
      const handle = await openFixture(client, 'empty.md', Buffer.from('', 'utf8'))
      const opened = await call(client, 'open_document', { path: 'empty.md' })
      expect(opened.structuredContent?.lineCount).toBe(1)
      const atEnd = await call(client, 'insert_content', { handle, text: 'hello' })
      expect(atEnd.isError).toBeFalsy()
      expect(atEnd.structuredContent?.lineCount).toBe(1)
      await call(client, 'save_document', { handle })
      expect((await readFile(join(root, 'empty.md'))).toString('utf8')).toBe('hello\n')

      const handle2 = await openFixture(client, 'empty2.md', Buffer.from('', 'utf8'))
      await call(client, 'insert_content', { handle: handle2, text: 'hello', at: -1 })
      await call(client, 'save_document', { handle: handle2 })
      expect((await readFile(join(root, 'empty2.md'))).toString('utf8')).toBe('hello\n')

      const handle3 = await openFixture(client, 'empty3.md', Buffer.from('', 'utf8'))
      const ops = await call(client, 'apply_ops', {
        handle: handle3,
        ops: [{ op: 'insertLines', after: 0, text: 'hello' }],
      })
      expect(ops.isError).toBeFalsy()
      await call(client, 'save_document', { handle: handle3 })
      expect((await readFile(join(root, 'empty3.md'))).toString('utf8')).toBe('hello\n')
    } finally {
      await close()
    }
  })
})

// Unit tests for the headless docx session: parse model, read formats,
// insert_content, apply_ops semantics (validation-forward, atomicity),
// byte-preservation, mtime fencing and path confinement.
import { mkdtemp, readFile, readdir, writeFile, rm, mkdir, stat, rename } from 'node:fs/promises'
import { existsSync, readdirSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

// the EPERM-fallback tests steer `link` through this vi.fn (pass-through to
// the real fs/promises link by default, like the sheets promote tests do for
// copyFile)
const { linkMock, actualLink, captureActual } = vi.hoisted(() => {
  let actual: ((...args: never[]) => Promise<void>) | undefined
  const passthrough = ((...args: never[]) => actual!(...args)) as (
    ...args: never[]
  ) => Promise<void>
  return {
    linkMock: vi.fn(passthrough),
    captureActual: (fn: (...args: never[]) => Promise<void>) => {
      actual = fn
    },
    actualLink: passthrough,
  }
})
vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  captureActual((...args: Parameters<typeof actual.link>) => actual.link(...args) as Promise<void>)
  return { ...actual, link: linkMock }
})

import {
  assertSaveTargetFree,
  DocxSession,
  FencingError,
  promoteNewFileExclusively,
} from '../src/docx/session.js'
import { resolveConfined, WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'
import type { Target } from '../src/docx/ops.js'
import { parseRestrictedHtml, blocksToHtml } from '../src/docx/html.js'
import { buildFixtureDocx } from './helpers/docx-fixture.js'

let root: string
let docPath: string

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-docx-'))
  docPath = join(root, 'report.docx')
  await writeFile(docPath, await buildFixtureDocx())
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

async function openSession(): Promise<DocxSession> {
  return DocxSession.open(docPath, root)
}

async function reparseSaved(path: string) {
  const { parseDocx } = await import('@airy-office/docx-engine')
  return parseDocx(new Uint8Array(await readFile(path)))
}

describe('docx session open/read', () => {
  it('opens a fixture and reports block/stats meta', async () => {
    const session = await openSession()
    const meta = session.meta()
    expect(meta.path).toBe(docPath)
    expect(meta.dirty).toBe(false)
    expect(meta.blockCount).toBe(7) // heading, para, 2 bullets, numbered item, table, tail
    expect(meta.wordCount).toBeGreaterThan(10)
  })

  it('renders the index|type|preview overview', async () => {
    const session = await openSession()
    const text = session.readDocument()
    expect(text).toContain('The document has 7 blocks')
    expect(text).toMatch(/^0\|h1\|Quarterly Report$/m)
    expect(text).toMatch(/^1\|p\|Revenue grew by 12 percent year over year\./m)
    expect(text).toMatch(/^2\|li\|First bullet item$/m)
    expect(text).toMatch(/^5\|table\|Region \| Sales \/ East \| 4200$/m)
    expect(text).toContain('Full-text stats:')
  })

  it('returns full restricted HTML for selected blocks', async () => {
    const session = await openSession()
    const text = session.readDocument({ blocks: [0, 1] })
    expect(text).toContain('<h1>Quarterly Report</h1>')
    expect(text).toContain('<p>Revenue grew by <strong>12 percent</strong> year over year.</p>')
  })

  it('groups consecutive list items into ul/ol on read', async () => {
    const session = await openSession()
    const text = session.readDocument({ range: { start: 2, end: 4 } })
    expect(text).toContain('<ul><li>First bullet item</li><li>Second bullet item</li></ul>')
    expect(text).toContain('<ol><li>Numbered step</li></ol>')
  })

  it('rejects out-of-range block selections', async () => {
    const session = await openSession()
    expect(() => session.readDocument({ blocks: [99] })).toThrow(/out of range/)
  })

  it('rejects an enormous read range fast, before allocating the index array', async () => {
    const session = await openSession()
    // before the span cap this loop built the whole index array first and
    // hung/OOMed the process; it must now fail fast with the cap message
    expect(() =>
      session.readDocument({ range: { start: 0, end: Number.MAX_SAFE_INTEGER } }),
    ).toThrow(/the cap is 10000 per read/)
    // a span just over the cap is refused; out-of-range counting still works
    expect(() => session.readDocument({ range: { start: 0, end: 10_000 } })).toThrow(
      /the cap is 10000 per read/,
    )
    expect(() => session.readDocument({ range: { start: 0, end: 99 } })).toThrow(/out of range/)
    // at or under the cap the range still reads normally
    const text = session.readDocument({ range: { start: 0, end: 6 } })
    expect(text).toContain('<h1>Quarterly Report</h1>')
  })
})

describe('restricted HTML parsing', () => {
  it('parses headings, inline marks, links and breaks', () => {
    const blocks = parseRestrictedHtml(
      '<h2>New <em>section</em></h2>' +
        '<p>Text with <strong>bold</strong>, <u>underline</u>, <s>strike</s>, ' +
        '<a href="https://example.com">a link</a> and a<br>line break.</p>',
    )
    expect(blocks).toHaveLength(2)
    expect(blocks[0]).toMatchObject({ type: 'heading', level: 2 })
    expect(blocks[1]).toMatchObject({ type: 'paragraph' })
    const runs = (blocks[1] as unknown as { runs: Array<Record<string, unknown>> }).runs
    expect(runs.some((r) => r.bold)).toBe(true)
    expect(runs.some((r) => r.underline)).toBe(true)
    expect(runs.some((r) => r.strike)).toBe(true)
    expect(runs.some((r) => r.link)).toBe(true)
    // <br> rides inside the run text as \n (the engine writes w:br for it)
    expect(runs.some((r) => typeof r.text === 'string' && r.text.includes('\n'))).toBe(true)
  })

  it('parses nested lists and tables', () => {
    const blocks = parseRestrictedHtml(
      '<ul><li>one<ol><li>nested</li></ol></li><li>two</li></ul>' +
        '<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>',
    )
    expect(blocks.map((b) => b.type)).toEqual(['listItem', 'listItem', 'listItem', 'table'])
    expect(blocks[1]).toMatchObject({ list: { kind: 'ordered', ilvl: 1 } })
    expect(blocks[3]).toHaveProperty('xml')
    expect((blocks[3] as { xml: string }).xml).toMatch(/^<w:tbl>/)
  })

  it('decodes entities and tolerates plain text / fences / unknown tags', () => {
    expect(parseRestrictedHtml('Fish &amp; Chips')).toEqual([
      { type: 'paragraph', runs: [{ text: 'Fish & Chips' }] },
    ])
    expect(parseRestrictedHtml('```html\n<p>hi</p>\n```')).toEqual([
      { type: 'paragraph', runs: [{ text: 'hi' }] },
    ])
    expect(parseRestrictedHtml('<span>kept</span>')).toEqual([
      { type: 'paragraph', runs: [{ text: 'kept' }] },
    ])
  })

  it('drops disallowed link schemes and keeps allowed ones (renderer href policy)', () => {
    const blocks = parseRestrictedHtml(
      '<p><a href="javascript:alert(1)">js</a> <a href="file:///etc/passwd">file</a> ' +
        '<a href="data:text/html,hi">data</a> <a href=" https://example.com/a ">https</a> ' +
        '<a href="mailto:a@b.c">mail</a> <a href="#frag">frag</a> ' +
        '<a href="docs/next.md">rel</a></p>',
    )
    const runs = (blocks[0] as unknown as { runs: Array<Record<string, unknown>> }).runs
    const links = runs
      .map((r) => (r.link as { href?: string } | undefined)?.href)
      .filter((h): h is string => h !== undefined)
    // disallowed schemes leave no link at all — their text survives as plain runs
    expect(links).not.toContain('javascript:alert(1)')
    expect(links).not.toContain('file:///etc/passwd')
    expect(links).not.toContain('data:text/html,hi')
    expect(runs.map((r) => r.text).join('')).toContain('js')
    expect(runs.map((r) => r.text).join('')).toContain('data')
    // allowed schemes survive, with surrounding whitespace trimmed
    expect(links).toContain('https://example.com/a')
    expect(links).toContain('mailto:a@b.c')
    expect(links).toContain('#frag')
    expect(links).toContain('docs/next.md')
  })
})

describe('insert_content', () => {
  it('inserts after the requested block and shifts indexes', async () => {
    const session = await openSession()
    const { inserted } = session.insertContent('<h3>Risks</h3><p>New paragraph.</p>', 1)
    expect(inserted).toBe(2)
    const text = session.readDocument({ blocks: [2, 3] })
    expect(text).toContain('<h3>Risks</h3>')
    expect(text).toContain('<p>New paragraph.</p>')
    expect(session.meta().blockCount).toBe(9)
    expect(session.meta().dirty).toBe(true)
  })

  it('inserts at the document start with at = -1 and at the end without at', async () => {
    const session = await openSession()
    session.insertContent('<p>top</p>', -1)
    expect(session.readDocument({ blocks: [0] })).toContain('<p>top</p>')
    session.insertContent('<p>bottom</p>')
    const count = session.meta().blockCount
    expect(session.readDocument({ blocks: [count - 1] })).toContain('<p>bottom</p>')
  })

  it('rejects HTML that parses into nothing', async () => {
    const session = await openSession()
    expect(() => session.insertContent('   ', 0)).toThrow(/did not parse/)
  })

  it('writes no external rel for dropped-scheme anchors; allowed hrefs persist', async () => {
    const session = await openSession()
    session.insertContent(
      '<p><a href="javascript:alert(1)">js</a> <a href="file:///etc/passwd">file</a> ' +
        '<a href="data:text/html,hi">data</a> <a href="https://example.com/ok">ok</a></p>',
      0,
    )
    await session.save()
    const zip = await JSZip.loadAsync(new Uint8Array(await readFile(docPath)))
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    // the engine writes every link run as an external rel target: the dropped
    // schemes must not appear there, the allowed https href must
    expect(rels).not.toContain('javascript:')
    expect(rels).not.toContain('file:')
    expect(rels).not.toContain('data:')
    expect(rels).toContain('https://example.com/ok')
    // degraded anchors still carry their text; only the allowed link reads back
    // as an anchor (at: 0 inserts after block 0, so the new block is index 1)
    const inserted = session.readDocument({ blocks: [1] })
    expect(inserted).toContain('js')
    expect(inserted).toContain('<a href="https://example.com/ok">ok</a>')
    expect(inserted).not.toContain('javascript:')
  })
})

describe('apply_ops', () => {
  it('validates the whole batch up front and applies nothing on error', async () => {
    const session = await openSession()
    expect(() =>
      session.applyOps([
        { op: 'findReplace', find: 'Revenue', replace: 'Profit' },
        { op: 'setFont', target: { nodeType: 'paragraph' }, bold: true, bogus: 1 },
      ]),
    ).toThrow(/unknown field/)
    // atomic: the valid first op must NOT have been applied
    expect(session.readDocument({ blocks: [1] })).toContain('Revenue')
    expect(session.meta().dirty).toBe(false)
  })

  it('rejects unknown op names and empty batches', async () => {
    const session = await openSession()
    expect(() => session.applyOps([{ op: 'nope' }])).toThrow(/unknown op/)
    expect(() => session.applyOps([])).toThrow(/non-empty/)
  })

  it('accepts both nodeType vocabularies with identical matching', async () => {
    // the renderer spellings (docHeading/docParagraph/docListItem) are
    // aliases of the headless canonical names, not separate block kinds: the
    // same batch in either vocabulary must produce the same document
    const canonical = await openSession()
    const aliased = await openSession()
    const canonicalOut = canonical.applyOps([
      { op: 'setFont', target: { nodeType: 'paragraph' }, bold: true },
      { op: 'setFont', target: { nodeType: 'heading' }, underline: true },
      { op: 'clearList', target: { nodeType: 'listItem' } },
    ])
    const aliasOut = aliased.applyOps([
      { op: 'setFont', target: { nodeType: 'docParagraph' }, bold: true },
      { op: 'setFont', target: { nodeType: 'docHeading' }, underline: true },
      { op: 'clearList', target: { nodeType: 'docListItem' } },
    ])
    expect(aliasOut.results.map((r) => r.matched)).toEqual(
      canonicalOut.results.map((r) => r.matched),
    )
    expect(aliasOut.results.map((r) => r.changed)).toEqual(
      canonicalOut.results.map((r) => r.changed),
    )
    // fixture: 2 paragraphs, 1 heading, 3 list items (2 bullets + 1 numbered)
    expect(canonicalOut.results.map((r) => r.matched)).toEqual([2, 1, 3])
    expect(aliased.readDocument()).toBe(canonical.readDocument())
  })

  it('unknown nodeType errors list the accepted spellings', async () => {
    const session = await openSession()
    // the wire accepts arbitrary strings; the validator must reject and explain
    const bogus = 'docImage' as unknown as Target['nodeType']
    expect(() =>
      session.applyOps([{ op: 'setFont', target: { nodeType: bogus }, bold: true }]),
    ).toThrow(
      /unknown nodeType "docImage".*heading\|docHeading.*paragraph\|docParagraph.*listItem\|docListItem.*image/,
    )
  })

  it('findReplace rewrites text inside runs', async () => {
    const session = await openSession()
    const { results } = session.applyOps([
      { op: 'findReplace', find: 'Revenue', replace: 'Profit' },
    ])
    expect(results[0]?.changed).toBe(1)
    expect(session.readDocument({ blocks: [1] })).toContain('Profit grew by')
  })

  it('case-insensitive findReplace and setMatchedFont survive the İ fold (BUG-1101)', async () => {
    const session = await openSession()
    // audit repro: İ (U+0130) lowercases to two code units, so lowered-index
    // slicing corrupted every run with Turkish text before it ("İstanbul
    // kXi" — tail eaten, replacement mid-word)
    session.insertContent('<p>İstanbul kelimesi ve not</p>')
    const { results } = session.applyOps([
      { op: 'findReplace', find: 'kelimesi', replace: 'sözcük', matchCase: false },
    ])
    expect(String(results[0]?.detail)).toContain('1 replacement(s)')
    expect(session.readDocument({ blocks: [7] })).toContain('İstanbul sözcük ve not')

    // occurrence styling slices runs on the same fold-safe ranges: the text
    // survives and only the occurrence run carries the style
    session.insertContent('<p>İstanbul kelimesi</p>')
    const styled = session.applyOps([
      { op: 'setMatchedFont', text: 'kelimesi', matchCase: false, bold: true },
    ])
    expect(String(styled.results[0]?.detail)).toContain('1 occurrence(s) styled')
    const html = session.readDocument({ blocks: [8] })
    expect(html).toContain('İstanbul <strong>kelimesi</strong>')
    expect(html).not.toContain('k<strong>')
  })

  it('setFont styles whole blocks; setMatchedFont styles occurrences only', async () => {
    const session = await openSession()
    session.applyOps([
      { op: 'setFont', target: { blockIndexes: [1] }, italic: true, color: '#1a73e8' },
    ])
    let html = session.readDocument({ blocks: [1] })
    expect(html).toContain('<em>')
    // occurrences-only styling keeps surrounding runs untouched
    const { results } = session.applyOps([{ op: 'setMatchedFont', text: 'bullet', bold: true }])
    expect(results[0]?.detail).toContain('2 occurrence(s)')
    html = session.readDocument({ blocks: [2] })
    expect(html).toContain('<strong>bullet</strong>')
  })

  it('setHeadingLevel and setParagraphFormat adjust structure/format', async () => {
    const session = await openSession()
    session.applyOps([
      { op: 'setHeadingLevel', target: { blockIndexes: [1] }, level: 2 },
      {
        op: 'setParagraphFormat',
        target: { nodeType: 'heading', headingLevel: 2 },
        align: 'center',
        spaceAfter: 240,
      },
    ])
    const text = session.readDocument({ blocks: [1] })
    expect(text).toContain('1|h2|')
    expect(text).toContain('<h2>Revenue grew by <strong>12 percent</strong> year over year.</h2>')
  })

  it('deleteBlocks and moveBlocks restructure the document', async () => {
    const session = await openSession()
    session.applyOps([{ op: 'deleteBlocks', target: { blockIndexes: [2, 3] } }])
    expect(session.meta().blockCount).toBe(5)
    // the numbered item now sits at index 2; move it right after the heading
    session.applyOps([{ op: 'moveBlocks', blockIndexes: [2], afterBlockIndex: 0 }])
    const text = session.readDocument()
    expect(text.split('\n')[2]).toBe('1|li|Numbered step')
  })

  it('setList converts paragraphs to list items with a real numbering id', async () => {
    const session = await openSession()
    session.applyOps([{ op: 'setList', target: { blockIndexes: [1] }, kind: 'number' }])
    const text = session.readDocument()
    expect(text.split('\n')[2]).toBe('1|li|Revenue grew by 12 percent year over year.')
    // the id must exist after save (restart or new def appended)
    await session.save(join(root, 'listed.docx'))
    const reparsed = await reparseSaved(join(root, 'listed.docx'))
    const block = reparsed.blocks.find((b) => b.type === 'listItem')
    expect(block?.list?.kind).toBe('ordered')
    expect(reparsed.numbering.has(block?.list?.numId ?? '')).toBe(true)
  })

  it('clearList converts list items back to paragraphs', async () => {
    const session = await openSession()
    session.applyOps([{ op: 'clearList', target: { nodeType: 'listItem' } }])
    const text = session.readDocument()
    expect(text).not.toMatch(/\|li\|/)
    expect(text).toMatch(/^2\|p\|First bullet item$/m)
  })

  it('dryRun reports but does not apply', async () => {
    const session = await openSession()
    const { summary } = session.applyOps(
      [{ op: 'findReplace', find: 'Revenue', replace: 'Profit' }],
      true,
    )
    expect(summary).toContain('changed 1')
    expect(session.readDocument({ blocks: [1] })).toContain('Revenue')
    expect(session.meta().dirty).toBe(false)
  })

  it('targets that match nothing change nothing (byte-safe no-op ops)', async () => {
    const session = await openSession()
    const { results } = session.applyOps([
      { op: 'setFont', target: { nodeType: 'image' }, bold: true },
      { op: 'findReplace', find: 'Absent Needle', replace: 'x' },
    ])
    expect(results[0]?.matched).toBe(0)
    expect(results[1]?.changed).toBe(0)
    expect(session.meta().dirty).toBe(false)
    // a no-op batch still saves the original bytes verbatim
    const result = await session.save()
    expect(result.unchanged).toBe(true)
  })
})

describe('save: byte preservation and fencing', () => {
  it('save with no edits returns the original bytes verbatim', async () => {
    const session = await openSession()
    const result = await session.save()
    expect(result.unchanged).toBe(true)
    expect(Buffer.compare(await readFile(docPath), await buildFixtureDocx())).toBe(0)
  })

  it('editing one block keeps every other zip entry and untouched paragraphs byte-identical', async () => {
    const session = await openSession()
    session.applyOps([{ op: 'findReplace', find: 'Revenue', replace: 'Profit' }])
    await session.save()
    const before = await JSZip.loadAsync(await buildFixtureDocx())
    const after = await JSZip.loadAsync(new Uint8Array(await readFile(docPath)))
    for (const path of ['word/styles.xml', 'word/numbering.xml', '_rels/.rels']) {
      expect(await after.file(path)!.async('string')).toBe(await before.file(path)!.async('string'))
    }
    const docBefore = await before.file('word/document.xml')!.async('string')
    const docAfter = await after.file('word/document.xml')!.async('string')
    // untouched paragraph keeps its exact bytes
    expect(docAfter).toContain('<w:r><w:t>End of report.</w:t></w:r>')
    expect(docAfter).toContain('Profit')
    expect(docBefore).toContain('Revenue')
    // the edited paragraph is the only body change: same number of paragraphs
    expect(docBefore.match(/<w:p[ >]/g)?.length).toBe(docAfter.match(/<w:p[ >]/g)?.length)
  })

  it('save-as writes a new file and leaves the original untouched', async () => {
    const session = await openSession()
    session.insertContent('<p>added</p>', 0)
    const target = join(root, 'nested', 'copy.docx')
    const result = await session.save(target)
    expect(result.path).toBe(target)
    const reparsed = await reparseSaved(target)
    const texts = reparsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts).toContain('added')
    // original on disk unchanged — byte-compare against a REBUILT fixture is
    // only deterministic because the builder pins every zip entry date (files
    // AND auto-created parent dirs; see helpers/docx-fixture.ts). Unpinned
    // dates flaked here on CI as TEST-721: jszip stamps entries with the
    // current time (2-second DOS granularity), so builds straddling a
    // 2-second boundary produced different bytes.
    expect(Buffer.compare(await readFile(docPath), await buildFixtureDocx())).toBe(0)
  })

  it('refuses save-as over an existing unrelated file unless overwrite is set', async () => {
    const session = await openSession()
    session.insertContent('<p>edit</p>', 0)
    const other = join(root, 'other.docx')
    await writeFile(other, "someone else's document")
    await expect(session.save(other)).rejects.toThrow(/already exists/)
    await expect(session.save(other)).rejects.toThrow(/overwrite: true/)
    // the refusal left the existing file untouched
    expect(await readFile(other, 'utf8')).toBe("someone else's document")
    // explicit consent replaces it
    await expect(session.save(other, 'docx', { overwrite: true })).resolves.toMatchObject({
      path: other,
    })
    expect(await readFile(other, 'utf8')).not.toBe("someone else's document")
    // the session's own opened file still saves without overwrite (fencing path)
    await expect(session.save(docPath)).resolves.toMatchObject({ path: docPath })
    // repeat save-as onto the session's own last output keeps working
    await expect(session.save(other)).resolves.toMatchObject({ path: other })
  })

  it('refuses to save when the file changed on disk since open', async () => {
    const session = await openSession()
    session.insertContent('<p>edit</p>', 0)
    // external writer: different size (and mtime)
    await writeFile(docPath, await buildFixtureDocx(), { flag: 'r+' })
    await new Promise((resolve) => setTimeout(resolve, 5))
    const other = await buildFixtureDocx()
    other[0] ^= 0xff
    await writeFile(docPath, other)
    await expect(session.save()).rejects.toThrow(FencingError)
    await expect(session.save()).rejects.toThrow(/changed on disk/)
  })

  it('fence refreshes after a successful save (chained saves work)', async () => {
    const session = await openSession()
    session.insertContent('<p>one</p>', 0)
    await session.save()
    session.insertContent('<p>two</p>', 0)
    await expect(session.save()).resolves.toMatchObject({ path: docPath })
  })

  it('refuses to save when the pinned workspace root was renamed away (BUG-1103)', async () => {
    // the session pins the root at open; renaming that directory mid-session
    // must fail the save clearly instead of mkdir-resurrecting the dead root
    const ws = join(root, 'ws')
    await mkdir(ws, { recursive: true })
    const wsDoc = join(ws, 'report.docx')
    await writeFile(wsDoc, await buildFixtureDocx())
    const session = await DocxSession.open(wsDoc, ws)
    session.insertContent('<p>edit</p>', 0)
    await rename(ws, join(root, 'ws2'))
    await expect(session.save(join('out.docx'))).rejects.toThrow(/no longer exists/)
    await expect(session.save(join('out.docx'))).rejects.toThrow(/Reopen the document/)
    // in-place save: the same stale-root refusal, not a drift fence error
    await expect(session.save()).rejects.toThrow(/no longer exists/)
    // the renamed-away directory was not resurrected
    expect(existsSync(ws)).toBe(false)
    expect(existsSync(join(root, 'ws2', 'out.docx'))).toBe(false)
  })

  it('cleans the tmp dotfile when the save fails after writing it (BUG-1111)', async () => {
    const session = await openSession()
    session.insertContent('<p>edit</p>', 0)
    // a non-empty DIRECTORY at the target: overwrite consent passes the
    // guard, but the promote's rename onto a non-empty directory fails
    // AFTER the `.<name>.airy-<uuid>` dotfile was written — the failure
    // must not orphan that dotfile next to the target forever
    const target = join(root, 'blocked.docx')
    await mkdir(target)
    await writeFile(join(target, 'keep'), 'contents')
    await expect(session.save(target, 'docx', { overwrite: true })).rejects.toThrow()
    const leftovers = (await readdir(root)).filter((name) => name.startsWith('.blocked.docx.airy-'))
    expect(leftovers).toEqual([])
  })
})

describe('assertSaveTargetFree (clobber guard)', () => {
  it('guards a converted session sibling like an explicit save-as', async () => {
    const sibling = join(root, 'legacy.docx')
    await writeFile(sibling, "someone else's document")
    // a converted session owns its temp .docx and the .doc origin, never the sibling
    const owned = [join(root, 'temp-converted.docx'), join(root, 'legacy.doc')]
    // default save onto the pre-existing sibling: refused, file untouched
    await expect(assertSaveTargetFree(sibling, owned, undefined)).rejects.toThrow(/already exists/)
    await expect(assertSaveTargetFree(sibling, owned, undefined)).rejects.toThrow(/overwrite: true/)
    expect(await readFile(sibling, 'utf8')).toBe("someone else's document")
    // explicit consent replaces it
    await expect(assertSaveTargetFree(sibling, owned, true)).resolves.toBeUndefined()
    // once saved, the sibling is the session's own output: repeat saves pass
    await expect(
      assertSaveTargetFree(sibling, [...owned, sibling], undefined),
    ).resolves.toBeUndefined()
    // a fresh (not yet existing) sibling target is allowed
    await expect(
      assertSaveTargetFree(join(root, 'fresh.docx'), owned, undefined),
    ).resolves.toBeUndefined()
  })

  it('lets native default saves (the opened file) through without overwrite', async () => {
    await expect(assertSaveTargetFree(docPath, [docPath], undefined)).resolves.toBeUndefined()
  })
})

describe('promoteNewFileExclusively (TOCTOU guard)', () => {
  it('refuses an existing target with the clobber error and cleans the temp', async () => {
    const tmp = join(root, '.new.docx.airy-test')
    const target = join(root, 'new.docx')
    await writeFile(tmp, 'new bytes')
    // a file created between the guard's stat and the write (the TOCTOU
    // window) must surface the actionable error, not be silently replaced
    await writeFile(target, 'created in the stat/write window')
    const error = await promoteNewFileExclusively(tmp, target).then(
      () => null,
      (e: Error) => e,
    )
    expect(error).toBeInstanceOf(Error)
    expect(error?.message).toContain('already exists')
    expect(error?.message).toContain('overwrite: true')
    expect(await readFile(target, 'utf8')).toBe('created in the stat/write window')
    await expect(stat(tmp)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('links a fresh target atomically and removes the temp', async () => {
    const tmp = join(root, '.fresh.docx.airy-test')
    const target = join(root, 'fresh.docx')
    await writeFile(tmp, 'fresh bytes')
    await expect(promoteNewFileExclusively(tmp, target)).resolves.toBeUndefined()
    expect(await readFile(target, 'utf8')).toBe('fresh bytes')
    await expect(stat(tmp)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('EPERM from link with an existing target still surfaces the clobber error', async () => {
    const tmp = join(root, '.eperm-exist.docx.airy-test')
    const target = join(root, 'eperm-exist.docx')
    await writeFile(tmp, 'new bytes')
    await writeFile(target, 'created in the stat/write window')
    // exFAT/FAT/network shares refuse hard links outright: the guard must
    // degrade to stat + clobber error, never a silent replace
    linkMock.mockImplementation(async () => {
      throw Object.assign(new Error('EPERM: operation not permitted, link'), { code: 'EPERM' })
    })
    try {
      const error = await promoteNewFileExclusively(tmp, target).then(
        () => null,
        (e: Error) => e,
      )
      expect(error?.message).toContain('already exists')
      expect(await readFile(target, 'utf8')).toBe('created in the stat/write window')
      await expect(stat(tmp)).rejects.toMatchObject({ code: 'ENOENT' })
    } finally {
      linkMock.mockImplementation(actualLink)
    }
  })

  it('EPERM from link with a missing target falls back to the atomic rename', async () => {
    const tmp = join(root, '.eperm-missing.docx.airy-test')
    const target = join(root, 'eperm-missing.docx')
    await writeFile(tmp, 'fallback bytes')
    linkMock.mockImplementation(async () => {
      throw Object.assign(new Error('EPERM: operation not permitted, link'), { code: 'EPERM' })
    })
    try {
      await expect(promoteNewFileExclusively(tmp, target)).resolves.toBeUndefined()
      expect(await readFile(target, 'utf8')).toBe('fallback bytes')
      await expect(stat(tmp)).rejects.toMatchObject({ code: 'ENOENT' })
    } finally {
      linkMock.mockImplementation(actualLink)
    }
  })
})

describe('path confinement', () => {
  it('rejects absolute paths outside the root', () => {
    expect(() => resolveConfined('/etc/passwd', root)).toThrow(/outside the workspace root/)
  })

  it('rejects traversal that escapes the root', () => {
    expect(() => resolveConfined('../../etc/passwd', root)).toThrow(/outside the workspace root/)
    expect(() => resolveConfined(join(root, '..', 'escape.docx'), root)).toThrow(
      /outside the workspace root/,
    )
  })

  it('allows paths inside the root and resolves relative ones against it', () => {
    expect(resolveConfined('a.docx', root)).toBe(join(root, 'a.docx'))
    expect(resolveConfined(join(root, 'sub', 'a.docx'), root)).toBe(join(root, 'sub', 'a.docx'))
  })

  it('rejects symlinks inside the root that point outside (real-path check)', async () => {
    // the outside target must exist: the real-path check only resolves when
    // every path component is present (a dangling link falls back lexical)
    const outsideDir = join(root, '..', 'airy-outside-target')
    await mkdir(outsideDir)
    await writeFile(join(outsideDir, 'secret.docx'), 'outside bytes')
    try {
      symlinkSync(outsideDir, join(root, 'escape'), 'dir')
    } catch (e) {
      // creating symlinks needs privileges on some platforms (Windows without
      // developer mode); confinement of the lexical path is still covered above
      const code = (e as NodeJS.ErrnoException).code
      if (code === 'EPERM' || code === 'EACCES') return
      throw e
    }
    try {
      // the link itself lives inside the root but resolves outside: rejected
      expect(() => resolveConfined(join(root, 'escape', 'secret.docx'), root)).toThrow(
        /outside the workspace root/,
      )
      // a relative hop through the link is rejected the same way
      expect(() => resolveConfined('escape/secret.docx', root)).toThrow(
        /outside the workspace root/,
      )
      // end to end: opening through the link is refused
      await expect(DocxSession.open(join(root, 'escape', 'secret.docx'), root)).rejects.toThrow(
        /outside the workspace root/,
      )
    } finally {
      await rm(outsideDir, { recursive: true, force: true })
    }
  })

  it('rejects a save-as to a NEW file through an out-pointing directory symlink', async () => {
    // the escape target directory must exist: the confinement check resolves
    // the deepest existing ancestor, so only a live link can pin the tail
    // outside the root
    const outsideDir = join(root, '..', 'airy-outside-new')
    await mkdir(outsideDir)
    try {
      symlinkSync(outsideDir, join(root, 'escape'), 'dir')
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code
      if (code === 'EPERM' || code === 'EACCES') return
      throw e
    }
    try {
      // the candidate file itself does not exist — the lexical fallback would
      // let it pass; the ancestor walk must catch the out-pointing link
      expect(() => resolveConfined('escape/new.docx', root)).toThrow(/outside the workspace root/)
      expect(() => resolveConfined(join(root, 'escape', 'new.docx'), root)).toThrow(
        /outside the workspace root/,
      )
      // end to end: a save-as through the link refuses and writes nothing outside
      const session = await openSession()
      await expect(session.save(join(root, 'escape', 'new.docx'))).rejects.toThrow(
        /outside the workspace root/,
      )
      expect(readdirSync(outsideDir)).toEqual([])
      await expect(readFile(join(outsideDir, 'new.docx'))).rejects.toThrow(/ENOENT/)
    } finally {
      await rm(outsideDir, { recursive: true, force: true })
    }
  })
  it('refuses a save-as onto an EXISTING symlink that points outside the root', async () => {
    // the target path itself is a file symlink to an outside document: the
    // real-path check resolves it (it exists) and refuses
    const outsideDir = join(root, '..', 'airy-outside-link-target')
    await mkdir(outsideDir)
    await writeFile(join(outsideDir, 'secret.docx'), 'outside bytes')
    try {
      symlinkSync(join(outsideDir, 'secret.docx'), join(root, 'linked.docx'), 'file')
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code
      if (code === 'EPERM' || code === 'EACCES') return
      throw e
    }
    try {
      expect(() => resolveConfined(join(root, 'linked.docx'), root)).toThrow(
        /outside the workspace root/,
      )
      // end to end: save-as onto the link refuses and never touches the target
      const session = await openSession()
      await expect(session.save(join(root, 'linked.docx'))).rejects.toThrow(
        /outside the workspace root/,
      )
      expect(await readFile(join(outsideDir, 'secret.docx'), 'utf8')).toBe('outside bytes')
    } finally {
      await rm(join(root, 'linked.docx'), { force: true })
      await rm(outsideDir, { recursive: true, force: true })
    }
  })

  it('refuses a fresh save-as through a multi-hop symlink chain (a -> b -> outside)', async () => {
    const outsideDir = join(root, '..', 'airy-outside-chain')
    await mkdir(outsideDir)
    try {
      symlinkSync(outsideDir, join(root, 'hop-b'), 'dir')
      symlinkSync(join(root, 'hop-b'), join(root, 'hop-a'), 'dir')
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code
      if (code === 'EPERM' || code === 'EACCES') return
      throw e
    }
    try {
      // the candidate file does not exist under the chain; the deepest
      // existing ancestor walk must follow BOTH hops out of the root
      expect(() => resolveConfined('hop-a/new.docx', root)).toThrow(/outside the workspace root/)
      expect(() => resolveConfined(join(root, 'hop-a', 'new.docx'), root)).toThrow(
        /outside the workspace root/,
      )
      const session = await openSession()
      await expect(session.save(join(root, 'hop-a', 'new.docx'))).rejects.toThrow(
        /outside the workspace root/,
      )
      expect(readdirSync(outsideDir)).toEqual([])
    } finally {
      await rm(join(root, 'hop-a'), { force: true })
      await rm(join(root, 'hop-b'), { force: true })
      await rm(outsideDir, { recursive: true, force: true })
    }
  })

  it('allows a fresh in-root save-as: nested dirs and deep missing chains', async () => {
    const session = await openSession()
    session.insertContent('<p>fresh</p>', 0)
    // an existing subdirectory with a new file inside stays allowed
    await mkdir(join(root, 'sub'), { recursive: true })
    const nested = await session.save(join(root, 'sub', 'new.docx'))
    expect(nested.path).toBe(join(root, 'sub', 'new.docx'))
    // a deep chain where only the root exists resolves back inside it (allowed)
    const deep = join(root, 'a', 'b', 'c', 'new.docx')
    expect(resolveConfined(deep, root)).toBe(deep)
    await expect(session.save(deep)).resolves.toMatchObject({ path: deep })
  })

  it('keeps symlinks usable when they resolve back inside the root', async () => {
    const inner = join(root, 'inner')
    await mkdir(inner)
    await writeFile(join(inner, 'real.docx'), await buildFixtureDocx())
    try {
      symlinkSync(inner, join(root, 'alias'), 'dir')
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code
      if (code === 'EPERM' || code === 'EACCES') return
      throw e
    }
    // opening through the in-root alias works and addresses the real file
    const session = await DocxSession.open(join(root, 'alias', 'real.docx'), root)
    expect(session.meta().blockCount).toBe(7)
  })

  it('keeps save-as working through a symlink-spelled AIRY_WORKSPACE_ROOT', async () => {
    const realRoot = await mkdtemp(join(tmpdir(), 'airy-docx-real-'))
    const link = join(root, 'rootlink')
    try {
      symlinkSync(realRoot, link, 'dir')
    } catch (e) {
      const code = (e as NodeJS.ErrnoException).code
      if (code === 'EPERM' || code === 'EACCES') {
        await rm(realRoot, { recursive: true, force: true })
        return
      }
      throw e
    }
    const previous = process.env[WORKSPACE_ROOT_ENV]
    process.env[WORKSPACE_ROOT_ENV] = link
    try {
      // the physical root and the env-spelled root differ; before the
      // deepest-ancestor fix every fresh save-as under the link was refused
      // (physical root vs lexical candidate failed the prefix check)
      await writeFile(join(realRoot, 'doc.docx'), await buildFixtureDocx())
      const session = await DocxSession.open('doc.docx')
      session.insertContent('<p>fresh</p>', 0)
      const saved = await session.save(join(link, 'nested', 'fresh.docx'))
      expect(saved.path).toBe(join(link, 'nested', 'fresh.docx'))
      // a path outside the link's target still refuses
      await expect(session.save(join(root, 'escape.docx'))).rejects.toThrow(
        /outside the workspace root/,
      )
    } finally {
      if (previous === undefined) delete process.env[WORKSPACE_ROOT_ENV]
      else process.env[WORKSPACE_ROOT_ENV] = previous
      await rm(realRoot, { recursive: true, force: true })
    }
  })

  it('folds case in the confinement compare on Windows-like platforms', () => {
    const realPlatform = process.platform
    try {
      Object.defineProperty(process, 'platform', { value: 'win32' })
      // drive-letter case differences must not reject a confined path
      expect(resolveConfined(join(root, 'a.docx'), root)).toBe(join(root, 'a.docx'))
      expect(resolveConfined(root.toUpperCase(), root.toUpperCase())).toBe(root.toUpperCase())
    } finally {
      Object.defineProperty(process, 'platform', { value: realPlatform })
    }
  })

  it('open and save enforce confinement end to end', async () => {
    const outside = join(root, '..', 'outside.docx')
    await expect(DocxSession.open(outside, root)).rejects.toThrow(/outside the workspace root/)
    const session = await openSession()
    await expect(session.save(outside)).rejects.toThrow(/outside the workspace root/)
    // a save into a fresh subdirectory works (mkdir -p semantics)
    await mkdir(join(root, 'out'), { recursive: true })
    await expect(session.save(join(root, 'out', 'x.docx'))).resolves.toBeTruthy()
  })

  it('rejects a non-docx file with a clear error', async () => {
    const bad = join(root, 'bad.docx')
    await writeFile(bad, 'this is not a zip')
    await expect(DocxSession.open(bad, root)).rejects.toThrow(/Cannot parse/)
  })
})

describe('read serialization of edited content', () => {
  it('round-trips inserted HTML through save and re-read', async () => {
    const session = await openSession()
    // h2 because the fixture's styles.xml only defines Heading1/Heading2 —
    // a heading level with no matching style would degrade to a paragraph
    session.insertContent(
      '<h2>Appendix</h2><ul><li>alpha</li><li>beta</li></ul>' +
        '<table><tr><th>K</th><th>V</th></tr><tr><td>one</td><td>1</td></tr></table>',
      6,
    )
    await session.save(join(root, 'round.docx'))
    const reparsed = await reparseSaved(join(root, 'round.docx'))
    const visible = reparsed.blocks.filter((b) => !b.hidden)
    const types = visible.map((b) => b.type)
    expect(types).toEqual([
      'heading',
      'paragraph',
      'listItem',
      'listItem',
      'listItem',
      'table',
      'paragraph',
      'heading',
      'listItem',
      'listItem',
      'table',
    ])
    const texts = visible.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts).toContain('Appendix')
    expect(texts).toContain('alpha')
    // the inserted table carries its header shading
    const insertedTable = visible.filter((b) => b.type === 'table')[1]
    expect(insertedTable?.table?.rows[0]?.[0]?.bold).toBe(true)
    // inserted list items join a real numbering definition
    const alpha = visible.find((b) => (b.runs ?? [])[0]?.text === 'alpha')
    expect(reparsed.numbering.has(alpha?.list?.numId ?? '')).toBe(true)
  })

  it('blocksToHtml emits escaped text', () => {
    const html = blocksToHtml([
      {
        type: 'paragraph',
        runs: [{ text: 'a < b & c' }],
      },
    ])
    expect(html).toBe('<p>a &lt; b &amp; c</p>')
  })
})

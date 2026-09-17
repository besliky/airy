// Unit tests for the headless docx session: parse model, read formats,
// insert_content, apply_ops semantics (validation-forward, atomicity),
// byte-preservation, mtime fencing and path confinement.
import { mkdtemp, readFile, writeFile, rm, mkdir } from 'node:fs/promises'
import { readdirSync, symlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { DocxSession, FencingError } from '../src/docx/session.js'
import { resolveConfined, WORKSPACE_ROOT_ENV } from '../src/docx/paths.js'
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

  it('findReplace rewrites text inside runs', async () => {
    const session = await openSession()
    const { results } = session.applyOps([
      { op: 'findReplace', find: 'Revenue', replace: 'Profit' },
    ])
    expect(results[0]?.changed).toBe(1)
    expect(session.readDocument({ blocks: [1] })).toContain('Profit grew by')
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
    // original on disk unchanged
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

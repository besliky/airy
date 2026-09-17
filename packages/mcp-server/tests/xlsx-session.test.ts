// Unit tests for the xlsx session: import routing (.xlsx vs .xls conversion),
// A1 parsing, read rendering, the cell-edit journal and the save/origin
// matrix — all against a stub sidecar IO and a mocked save module (the real
// gateway save needs the sidecar binary; covered by the integration file).
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

// capture of every saveWorkbookViaSidecar call; the mock writes the target so
// the session's post-save stat works
const saveCalls: Array<Record<string, unknown>> = []
vi.mock('../src/xlsx/save.js', () => ({
  saveWorkbookViaSidecar: vi.fn(async (request: Record<string, unknown>) => {
    saveCalls.push(request)
    await writeFile(String(request.targetPath), 'saved-xlsx-bytes')
    return { touchedEntries: ['xl/worksheets/sheet1.xml'], removedEntries: [], addedEntries: [] }
  }),
}))

// LibreOffice is never present in these unit tests: the origin-export error
// paths assert the actionable "install LibreOffice" message
vi.mock('../src/import/soffice.js', () => ({
  SOFFICE_FILTERS: { docx: 'MS Word 2007 XML', doc: 'MS Word 97', odt: 'writer8', ods: 'calc8' },
  findSoffice: vi.fn(async () => null),
  sofficeMissingError: (reason: string) => new Error(`${reason} LibreOffice is not installed.`),
  convertViaSoffice: vi.fn(async () => {
    throw new Error('not available in unit tests')
  }),
}))

import { XlsxSession, parseA1Range } from '../src/xlsx/session.js'
import { FencingError } from '../src/docx/session.js'
import { makeStubIo } from './helpers/stub-sidecar.js'

let root: string

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-xlsx-'))
  saveCalls.length = 0
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

describe('A1 notation', () => {
  it('parses single cells and ranges, normalizing reversed corners', () => {
    expect(parseA1Range('B2')).toEqual({ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 })
    expect(parseA1Range('A1:C10')).toEqual({
      startRow: 0,
      endRow: 9,
      startColumn: 0,
      endColumn: 2,
    })
    expect(parseA1Range('C10:A1')).toEqual({
      startRow: 0,
      endRow: 9,
      startColumn: 0,
      endColumn: 2,
    })
    expect(parseA1Range('$AA$5')).toEqual({
      startRow: 4,
      endRow: 4,
      startColumn: 26,
      endColumn: 26,
    })
  })

  it('rejects malformed specs', () => {
    expect(() => parseA1Range('A1:')).toThrow(/Invalid A1-style range/)
    expect(() => parseA1Range('banana')).toThrow(/Invalid A1-style range/)
  })
})

describe('XlsxSession.open', () => {
  it('opens .xlsx natively: meta is editable, unconverted, warning-free', async () => {
    const bookPath = join(root, 'book.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const io = makeStubIo()
    const session = await XlsxSession.open(bookPath, root, io)
    const meta = session.meta()
    expect(meta.kind).toBe('xlsx')
    expect(meta.format).toBe('xlsx')
    expect(meta.converted).toBe(false)
    expect(meta.editable).toBe(true)
    expect(meta.warnings).toEqual([])
    expect(meta.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1', 'Data'])
    expect(meta.path).toBe(bookPath)
    expect(io.calls.open).toEqual([bookPath])
  })

  it('imports .xls via convert_workbook with a styles-lost warning and temp backing', async () => {
    const legacyPath = join(root, 'legacy.xls')
    await writeFile(legacyPath, 'legacy-bytes')
    const io = makeStubIo()
    const session = await XlsxSession.open(legacyPath, root, io)
    const meta = session.meta()
    expect(meta.format).toBe('xls')
    expect(meta.converted).toBe(true)
    expect(meta.path).toBe(legacyPath) // user-facing path stays the origin
    expect(meta.warnings[0]).toMatch(/lost on import/)
    expect(io.calls.convert).toHaveLength(1)
    // the conversion target became the sidecar-opened backing file
    const convertedTarget = String(io.calls.convert[0]).split('->')[1]
    expect(io.calls.open).toEqual([convertedTarget])
    await session.close()
  })

  it('rejects unsupported workbook extensions before touching the sidecar', async () => {
    const io = makeStubIo()
    await expect(XlsxSession.open(join(root, 'book.xlsb'), root, io)).rejects.toThrow(
      /Unsupported workbook extension/,
    )
    expect(io.calls.open).toEqual([])
  })

  it('cleans the temp conversion dir when open fails after conversion', async () => {
    const legacyPath = join(root, 'broken.xls')
    await writeFile(legacyPath, 'legacy-bytes')
    const io = makeStubIo({ openError: new Error('workbook_error: bad zip') })
    await expect(XlsxSession.open(legacyPath, root, io)).rejects.toThrow(/bad zip/)
    expect(io.calls.convert).toHaveLength(1)
    const convertedTarget = String(io.calls.convert[0]).split('->')[1]
    const { access } = await import('node:fs/promises')
    await expect(access(String(convertedTarget))).rejects.toThrow()
  })
})

describe('XlsxSession.readWorkbook', () => {
  it('renders the sheet overview without options', async () => {
    const bookPath = join(root, 'book.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const session = await XlsxSession.open(bookPath, root, makeStubIo())
    const text = await session.readWorkbook()
    expect(text).toContain('The workbook has 2 sheet(s)')
    expect(text).toMatch(/0\|Sheet1\|sheet-0\|100 x 26/)
    expect(text).toMatch(/1\|Data\|sheet-1\|50 x 10/)
    expect(text).toContain('read_workbook')
  })

  it('renders values and formulas as a pipe table', async () => {
    const bookPath = join(root, 'book.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const io = makeStubIo({
      cells: [
        { row: 0, column: 0, value: 'Region' },
        { row: 0, column: 1, value: 'Sales' },
        { row: 1, column: 0, value: 'East' },
        { row: 1, column: 1, value: 42 },
        { row: 2, column: 1, value: 63, formula: 'SUM(B2:B2)' },
        { row: 2, column: 0, value: true },
      ],
    })
    const session = await XlsxSession.open(bookPath, root, io)
    const text = await session.readWorkbook({ sheet: 'Sheet1', range: 'A1:B3' })
    expect(text).toContain('Sheet "Sheet1", range A1:B3:')
    expect(text).toContain('| |A|B|')
    expect(text).toContain('|1|Region|Sales|')
    expect(text).toContain('|2|East|42|')
    expect(text).toContain('|3|TRUE|=SUM(B2:B2) (63)|')
  })

  it('resolves sheets by index and reports unknown sheets clearly', async () => {
    const bookPath = join(root, 'book.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const session = await XlsxSession.open(bookPath, root, makeStubIo())
    await expect(session.readWorkbook({ sheet: 1, range: 'A1:A1' })).resolves.toContain('Data')
    await expect(session.readWorkbook({ sheet: 'Nope' })).rejects.toThrow(/No sheet "Nope"/)
    await expect(session.readWorkbook({ range: 'A1:A1' })).rejects.toThrow(
      /range requires the sheet/,
    )
  })

  it('refuses oversized ranges', async () => {
    const bookPath = join(root, 'book.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const io = makeStubIo({
      sheets: [{ id: 'sheet-0', name: 'Big', rowCount: 500, columnCount: 100 }],
    })
    const session = await XlsxSession.open(bookPath, root, io)
    await expect(session.readWorkbook({ sheet: 'Big', range: 'A1:CV500' })).rejects.toThrow(
      /Split it into smaller reads/,
    )
  })
})

describe('XlsxSession journal + save matrix', () => {
  async function nativeSession() {
    const bookPath = join(root, 'book.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const io = makeStubIo()
    return { session: await XlsxSession.open(bookPath, root, io), io, bookPath }
  }

  it('journals cell edits with last-write-wins dedupe and = prefixes for formulas', async () => {
    const { session } = await nativeSession()
    session.setCells({
      sheet: 'Sheet1',
      cells: [
        { ref: 'A1', value: 'first' },
        { ref: 'A1', value: 'second' },
        { ref: 'B2', formula: 'SUM(A1:A1)' },
      ],
    })
    await session.save(join(root, 'out.xlsx'))
    const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
    expect(edits).toHaveLength(2)
    expect(edits[0]).toMatchObject({
      sheetName: 'Sheet1',
      row: 0,
      column: 0,
      cell: { value: 'second' },
    })
    expect(edits[1]).toMatchObject({
      row: 1,
      column: 1,
      cell: { value: '', formula: '=SUM(A1:A1)' },
    })
    // journal flushed after save
    expect(session.meta().dirty).toBe(false)
  })

  it('saves in place by default, reopens the sidecar session afterwards', async () => {
    const { session, io } = await nativeSession()
    const result = await session.save()
    expect(result.path).toBe(join(root, 'book.xlsx'))
    expect(result.format).toBe('xlsx')
    expect(result.unchanged).toBe(true) // no edits journaled
    // in-place save closes and reopens the sidecar session for fresh reads
    expect(io.calls.close).toHaveLength(1)
    expect(io.calls.open).toEqual([join(root, 'book.xlsx'), join(root, 'book.xlsx')])
  })

  it('refuses an in-place save when the backing file changed on disk since open', async () => {
    const { session } = await nativeSession()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'edit' }] })
    // external writer rewrites the backing file between open and save
    await writeFile(join(root, 'book.xlsx'), 'externally rewritten bytes')
    await expect(session.save()).rejects.toThrow(FencingError)
    await expect(session.save()).rejects.toThrow(/changed on disk/)
    // the refusal left the external writer's file untouched
    expect(await readFile(join(root, 'book.xlsx'), 'utf8')).toBe('externally rewritten bytes')
  })

  it('refreshes the fence after a successful in-place save (chained saves work)', async () => {
    const { session } = await nativeSession()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'one' }] })
    await session.save()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'two' }] })
    await expect(session.save()).resolves.toMatchObject({ path: join(root, 'book.xlsx') })
  })

  it('refuses save-as over an existing unrelated file unless overwrite is set', async () => {
    const { session } = await nativeSession()
    const other = join(root, 'other.xlsx')
    await writeFile(other, 'unrelated workbook bytes')
    await expect(session.save(other)).rejects.toThrow(/already exists/)
    await expect(session.save(other)).rejects.toThrow(/overwrite: true/)
    // the refusal left the existing file untouched
    expect(await readFile(other, 'utf8')).toBe('unrelated workbook bytes')
    // explicit consent replaces it
    await expect(session.save(other, 'xlsx', { overwrite: true })).resolves.toMatchObject({
      path: other,
      format: 'xlsx',
    })
    expect(await readFile(other, 'utf8')).toBe('saved-xlsx-bytes')
    // the session's own backing file still saves without overwrite (fencing path)
    await expect(session.save()).resolves.toMatchObject({ path: join(root, 'book.xlsx') })
    // repeat save-as onto the session's own last output keeps working
    await expect(session.save(other)).resolves.toMatchObject({ path: other })
  })

  it('defaults .xls imports to a sibling .xlsx and leaves the origin untouched', async () => {
    const legacyPath = join(root, 'legacy.xls')
    await writeFile(legacyPath, 'legacy-bytes')
    const session = await XlsxSession.open(legacyPath, root, makeStubIo())
    const result = await session.save()
    expect(result.path).toBe(join(root, 'legacy.xlsx'))
    expect(result.warnings[0]).toMatch(/original .*\.xls.* left untouched/)
  })

  it('refuses a converted session default save onto a pre-existing sibling', async () => {
    const legacyPath = join(root, 'legacy.xls')
    await writeFile(legacyPath, 'legacy-bytes')
    const session = await XlsxSession.open(legacyPath, root, makeStubIo())
    const sibling = join(root, 'legacy.xlsx')
    await writeFile(sibling, 'pre-existing sibling bytes')
    // the sibling is not a file the session opened or saved: the default
    // save must not silently clobber it
    await expect(session.save()).rejects.toThrow(/already exists/)
    await expect(session.save()).rejects.toThrow(/overwrite: true/)
    expect(await readFile(sibling, 'utf8')).toBe('pre-existing sibling bytes')
    // explicit consent replaces it
    await expect(session.save(undefined, 'xlsx', { overwrite: true })).resolves.toMatchObject({
      path: sibling,
    })
    expect(await readFile(sibling, 'utf8')).toBe('saved-xlsx-bytes')
    // after the first save the sibling is the session's own output
    await expect(session.save()).resolves.toMatchObject({ path: sibling })
  })

  it('asks the gateway for an exclusive promote only on guarded fresh targets', async () => {
    const { session } = await nativeSession()
    // a fresh save-as target promotes exclusively (TOCTOU-safe)
    await session.save(join(root, 'fresh.xlsx'))
    expect(saveCalls[0]?.exclusiveTarget).toBe(true)
    // in-place saves replace the session's own backing file by intent
    saveCalls.length = 0
    await session.save()
    expect(saveCalls[0]?.exclusiveTarget).toBe(false)
    // explicit overwrite consent replaces by intent
    saveCalls.length = 0
    const other = join(root, 'other.xlsx')
    await writeFile(other, 'unrelated bytes')
    await session.save(other, 'xlsx', { overwrite: true })
    expect(saveCalls[0]?.exclusiveTarget).toBe(false)
  })

  it('format origin: .xls refuses with the save-as-.xlsx cascade', async () => {
    const legacyPath = join(root, 'legacy.xls')
    await writeFile(legacyPath, 'legacy-bytes')
    const session = await XlsxSession.open(legacyPath, root, makeStubIo())
    await expect(session.save(undefined, 'origin')).rejects.toThrow(
      /Writing the legacy \.xls format is not supported/,
    )
  })

  it('format origin: native sessions reject, .ods needs LibreOffice', async () => {
    const { session } = await nativeSession()
    await expect(session.save(undefined, 'origin')).rejects.toThrow(
      /only valid for sessions imported from \.xls\/\.ods/,
    )
    const odsPath = join(root, 'book.ods')
    await writeFile(odsPath, 'ods-bytes')
    const odsSession = await XlsxSession.open(odsPath, root, makeStubIo())
    await expect(odsSession.save(undefined, 'origin')).rejects.toThrow(
      /LibreOffice is not installed/,
    )
  })

  it('close() closes the sidecar session and removes the import temp dir', async () => {
    const legacyPath = join(root, 'legacy.xls')
    await writeFile(legacyPath, 'legacy-bytes')
    const io = makeStubIo()
    const session = await XlsxSession.open(legacyPath, root, io)
    const backing = String(io.calls.convert[0]).split('->')[1]
    const cleaned = await session.close()
    expect(io.calls.close).toHaveLength(1)
    expect(cleaned).toHaveLength(1)
    const { access } = await import('node:fs/promises')
    await expect(access(backing)).rejects.toThrow()
  })
})

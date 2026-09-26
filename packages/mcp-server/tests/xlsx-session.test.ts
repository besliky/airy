// Unit tests for the xlsx session: import routing (.xlsx vs .xls conversion),
// A1 parsing, read rendering, the cell-edit journal and the save/origin
// matrix — all against a stub sidecar IO and a mocked save module (the real
// gateway save needs the sidecar binary; covered by the integration file).
import { existsSync } from 'node:fs'
import { mkdtemp, mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

// capture of every saveWorkbookViaSidecar call; the mock writes the target so
// the session's post-save stat works
const saveCalls: Array<Record<string, unknown>> = []
vi.mock('../src/xlsx/save.js', () => ({
  // the session's save catch tests its rejections against this class, so the
  // mock must carry it (never instantiated in unit tests)
  SaveTargetExistsError: class extends Error {},
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

import { XlsxSession, setRecalcIndexWaitForTests } from '../src/xlsx/session.js'
import { saveWorkbookViaSidecar } from '../src/xlsx/save.js'
import { MAX_SHEET_COLUMNS, MAX_SHEET_ROWS, parseA1Range } from '../src/xlsx/refs.js'
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

  it('rejects syntax-valid refs that address no cell (BUG-1632)', () => {
    // "A0" parses to 0-based row -1: the audit repro that used to journal
    expect(() => parseA1Range('A0')).toThrow(
      /Ref "A0" addresses no cell: rows and columns start at 1 in A1 notation/,
    )
    expect(() => parseA1Range('A00')).toThrow(/addresses no cell/)
    expect(() => parseA1Range('$A$0')).toThrow(/addresses no cell/)
    expect(() => parseA1Range('A0:B2')).toThrow(/addresses no cell/)
    // past the SpreadsheetML grid there is no cell either
    expect(() => parseA1Range('A1048577')).toThrow(/past the sheet's last cell/)
    expect(() => parseA1Range('XFE1')).toThrow(/past the sheet's last cell/)
  })

  it('accepts the grid boundary refs (A1 and the XFD1048576 corner)', () => {
    expect(parseA1Range('A1')).toEqual({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })
    expect(parseA1Range('XFD1048576')).toEqual({
      startRow: MAX_SHEET_ROWS - 1,
      endRow: MAX_SHEET_ROWS - 1,
      startColumn: MAX_SHEET_COLUMNS - 1,
      endColumn: MAX_SHEET_COLUMNS - 1,
    })
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

describe('read-after-write journal overlay (BUG-1503)', () => {
  function salesIo() {
    return makeStubIo({
      sheets: [{ id: 'sheet-0', name: 'Sales', rowCount: 11, columnCount: 4 }],
      cells: [
        { row: 0, column: 0, value: 'Region' },
        { row: 0, column: 1, value: 'Q1' },
        { row: 1, column: 0, value: 'North' },
        { row: 1, column: 1, value: 100 },
      ],
    })
  }

  async function openSales(io = salesIo()) {
    const bookPath = join(root, 'sales.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    return { session: await XlsxSession.open(bookPath, root, io), io, bookPath }
  }

  it('reads show journaled edits before the save, not the stale file', async () => {
    // audit repro: apply_workbook_ops A1/B2 -> read_workbook A1:B2 used to
    // return the on-disk "Region|Q1|North|100" although the edits were journaled
    const { session } = await openSales()
    session.setCells({
      sheet: 'Sales',
      cells: [
        { ref: 'A1', value: 'EDITED-HEADER' },
        { ref: 'B2', value: 999 },
      ],
    })
    const text = await session.readWorkbook({ sheet: 'Sales', range: 'A1:B2' })
    expect(text).toContain('|1|EDITED-HEADER|Q1|')
    expect(text).toContain('|2|North|999|')
    expect(text).toContain('2 pending edit(s)')

    // a null clears the cell; a formula renders without a cached value
    session.setCells({
      sheet: 'Sales',
      cells: [
        { ref: 'A1', value: null },
        { ref: 'B1', formula: 'SUM(B2:B2)' },
      ],
    })
    const updated = await session.readWorkbook({ sheet: 'Sales', range: 'A1:B2' })
    expect(updated).toContain('|1||=SUM(B2:B2)|')
    expect(updated).toContain('|2|North|999|')
  })

  it('style-only edits do not change the value table', async () => {
    const { session } = await openSales()
    session.setCells({ sheet: 'Sales', cells: [{ ref: 'A1', style: { bold: true } }] })
    const text = await session.readWorkbook({ sheet: 'Sales', range: 'A1:B2' })
    expect(text).toContain('|1|Region|Q1|')
    expect(text).not.toContain('pending edit')
  })

  it('journaled cells beyond the used range are readable', async () => {
    // audit repro: edits into column E (beyond 11x4) made read E1:E4 fail
    // with "Range is outside sheet" until save
    const io = salesIo()
    const { session } = await openSales(io)
    session.setCells({
      sheet: 'Sales',
      cells: [
        { ref: 'E1', value: 'new col' },
        { ref: 'E2', value: 7 },
        { ref: 'E3', value: 8 },
        { ref: 'E4', value: 9 },
      ],
    })
    // the overview reports the journal-grown area
    const overview = await session.readWorkbook()
    expect(overview).toMatch(/0\|Sales\|sheet-0\|11 x 5/)
    const text = await session.readWorkbook({ sheet: 'Sales', range: 'E1:E4' })
    expect(text).toContain('|1|new col|')
    expect(text).toContain('|2|7|')
    expect(text).toContain('|4|9|')
    expect(text).toContain('4 pending edit(s)')
    // the sidecar only ever sees the on-disk 11x4 intersection (the real
    // binary refuses ranges past the used area)
    expect(io.readRanges.every((spec) => !/ x 4\.\.\d+/.test(spec))).toBe(true)
    // beyond the journal-extended area still refuses, with the grown dims
    await expect(session.readWorkbook({ sheet: 'Sales', range: 'F1:F1' })).rejects.toThrow(
      /Range is outside sheet "Sales" \(11 rows x 5 columns\)/,
    )
  })

  it('a journaled cell on an empty sheet makes the corner read non-empty', async () => {
    const io = makeStubIo({
      sheets: [{ id: 'sheet-0', name: 'Blank', rowCount: 0, columnCount: 0 }],
    })
    const { session } = await openSales(io)
    const before = await session.readWorkbook({ sheet: 'Blank' })
    expect(before).toContain('is empty')
    session.setCells({ sheet: 'Blank', cells: [{ ref: 'A1', value: 'hello' }] })
    const after = await session.readWorkbook({ sheet: 'Blank' })
    expect(after).toContain('|1|hello|')
    expect(after).toContain('1 pending edit(s)')
  })

  it('reads after save reflect the persisted file again (journal flushed)', async () => {
    const { session } = await openSales()
    session.setCells({ sheet: 'Sales', cells: [{ ref: 'A1', value: 'EDITED' }] })
    await session.save()
    // journal flushed: the stub sidecar still serves the original canned
    // cells (the mocked save never wrote them), and no overlay note appears
    const text = await session.readWorkbook({ sheet: 'Sales', range: 'A1:A1' })
    expect(text).toContain('|1|Region|')
    expect(text).not.toContain('pending edit')
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

  it('refuses the "A0" op without journaling; the session stays alive and saves (BUG-1632)', async () => {
    const { session } = await nativeSession()
    // audit repro: {"ref":"A0","value":1} used to journal row -1 ("Journaled
    // 1 cell edit(s)") and every later save then died with "Invalid cell
    // coordinates: -1,0", forcing a close that lost all unsaved edits
    expect(() => session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A0', value: 1 }] })).toThrow(
      /Ref "A0" addresses no cell/,
    )
    // the journal took nothing, so the next save works and writes nothing
    expect(session.meta().dirty).toBe(false)
    await expect(session.save()).resolves.toMatchObject({ unchanged: true })
    expect(saveCalls[0]?.edits).toEqual([])
  })

  it('a batch containing a bad ref journals nothing; earlier valid edits survive (BUG-1632)', async () => {
    const { session } = await nativeSession()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'good' }] })
    expect(() =>
      session.setCells({
        sheet: 'Sheet1',
        cells: [
          { ref: 'B2', value: 1 },
          { ref: 'A0', value: 2 },
        ],
      }),
    ).toThrow(/Ref "A0" addresses no cell/)
    // only the earlier good edit reaches the save — not the refused batch
    await expect(session.save()).resolves.toMatchObject({ unchanged: false })
    const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
    expect(edits).toHaveLength(1)
    expect(edits[0]).toMatchObject({
      sheetName: 'Sheet1',
      row: 0,
      column: 0,
      cell: { value: 'good' },
    })
  })

  it('boundary refs journal cleanly (A1 and the XFD1048576 corner)', async () => {
    const { session } = await nativeSession()
    const { journaled, merged } = session.setCells({
      sheet: 'Sheet1',
      cells: [
        { ref: 'A1', value: 'top-left' },
        { ref: 'XFD1048576', value: 'bottom-right' },
      ],
    })
    expect(journaled).toBe(2)
    expect(merged).toBe(2)
    expect(session.meta().dirty).toBe(true)
    await expect(session.save()).resolves.toMatchObject({ unchanged: false })
    const edits = saveCalls[0]?.edits as Array<Record<string, unknown>>
    expect(edits[1]).toMatchObject({
      row: MAX_SHEET_ROWS - 1,
      column: MAX_SHEET_COLUMNS - 1,
      cell: { value: 'bottom-right' },
    })
  })

  it('reads refuse exactly the refs the write path refuses (one shared validator)', async () => {
    const { session } = await nativeSession()
    let writeError: Error | undefined
    try {
      session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A0', value: 1 }] })
    } catch (e) {
      writeError = e as Error
    }
    expect(writeError).toBeDefined()
    // the read path produces the very message the write path refused with —
    // both run the same parse/validate in refs.ts (the divergence that let
    // "A0" journal while reads said "Range is outside sheet" is closed)
    await expect(session.readWorkbook({ sheet: 'Sheet1', range: 'A0' })).rejects.toThrow(
      writeError!.message,
    )
    await expect(session.readWorkbook({ sheet: 'Sheet1', range: 'A1048577' })).rejects.toThrow(
      /past the sheet's last cell/,
    )
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

  it('names a read-only workspace when the sidecar write refuses (UX-1690)', async () => {
    const { session } = await nativeSession()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'edit' }] })
    // the save writes through the Rust sidecar, so the refusal arrives as
    // plain error text without a Node errno `.code` (the audit repro: a 0555
    // workspace answered with the raw errno / sidecar io text)
    vi.mocked(saveWorkbookViaSidecar).mockImplementationOnce(async () => {
      throw new Error('failed to create target: Permission denied (os error 13)')
    })
    const outcome = session.save()
    await expect(outcome).rejects.toThrow(/Cannot write ".*": permission denied/)
    await expect(outcome).rejects.toThrow(/may be read-only/)
    // the sidecar's raw errno text no longer leaks
    await expect(outcome).rejects.not.toThrow(/os error 13/)
    // the failed save kept the journal so the edits can be retried elsewhere
    expect(session.meta().dirty).toBe(true)
  })

  it('refreshes the fence after a successful in-place save (chained saves work)', async () => {
    const { session } = await nativeSession()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'one' }] })
    await session.save()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'two' }] })
    await expect(session.save()).resolves.toMatchObject({ path: join(root, 'book.xlsx') })
  })

  it('refuses to save when the pinned workspace root was renamed away (BUG-1103 parity)', async () => {
    // the session pins the root at open; renaming that directory mid-session
    // must fail with the documented stale-root refusal (like docx/line/slides
    // sessions), not a raw ENOENT from the save gateway
    const ws = join(root, 'ws')
    await mkdir(ws, { recursive: true })
    const wsBook = join(ws, 'book.xlsx')
    await writeFile(wsBook, 'stub-xlsx-bytes')
    const session = await XlsxSession.open(wsBook, ws, makeStubIo())
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 'edit' }] })
    await rename(ws, join(root, 'ws2'))
    // save-as, in-place and origin-export all refuse with the same error
    await expect(session.save(join('out.xlsx'))).rejects.toThrow(/no longer exists/)
    await expect(session.save()).rejects.toThrow(/no longer exists/)
    await expect(session.save(undefined, 'origin')).rejects.toThrow(/no longer exists/)
    await expect(session.save()).rejects.toThrow(/Reopen the document/)
    // the gateway was never reached and the dead root stayed dead
    expect(saveCalls).toHaveLength(0)
    expect(existsSync(ws)).toBe(false)
    expect(existsSync(join(root, 'out.xlsx'))).toBe(false)
    await session.close()
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

// BUG-1761 (FM-7, the MCP twin of #230's recalc overlay): an edited save used
// to keep the file's own formula caches — a formula-only book saved with no
// <v> at all, a cached book kept stale numbers after its inputs changed, and
// script readers (openpyxl data_only, pandas) silently read wrong values. The
// save now evaluates the journal through the sidecar's recalc engine and
// overlays the fresh results into <v>, degrading honestly when the engine is
// unavailable.
describe('save refreshes formula caches (BUG-1761)', () => {
  interface FormulaHarness {
    session: XlsxSession
    io: ReturnType<typeof makeStubIo>
    bookPath: string
  }

  async function formulaSession(
    options: Parameters<typeof makeStubIo>[0] = {},
  ): Promise<FormulaHarness> {
    const bookPath = join(root, 'formula-wb.xlsx')
    await writeFile(bookPath, 'stub-xlsx-bytes')
    const io = makeStubIo({
      formulaCells: [
        { sheetId: 'sheet-0', row: 2, column: 0, value: 5 },
        { sheetId: 'sheet-0', row: 0, column: 1, value: 50 },
      ],
      recalcCells: [
        { sheet: 'Sheet1', row: 2, column: 0, number: 103, isFormula: true },
        { sheet: 'Sheet1', row: 0, column: 1, number: 1030, isFormula: true },
      ],
      ...options,
    })
    const session = await XlsxSession.open(bookPath, root, io)
    return { session, io, bookPath }
  }

  it('overlays recalculated values (not the file caches) into the save', async () => {
    const { session, io } = await formulaSession()
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 100 }] })
    const saved = await session.save(join(root, 'out.xlsx'))
    expect(saved.warnings).toEqual([])

    // the recalc request carried the journal edit as user input and read the
    // formula cells 1x1 (A3 = row 2, col 0)
    expect(io.recalcRequests).toHaveLength(1)
    expect(io.recalcRequests[0]?.edits).toEqual([
      { sheet: 'Sheet1', row: 0, column: 0, input: '100' },
    ])
    expect(io.recalcRequests[0]?.reads).toEqual([
      { sheet: 'Sheet1', range: { startRow: 2, endRow: 2 } },
      { sheet: 'Sheet1', range: { startRow: 0, endRow: 0 } },
    ])

    // the save gateway receives the fresh engine values — NOT the file's own
    // stale caches (5/50); both non-journaled formula cells ride the overlay
    expect(saveCalls[0]?.formulaValues).toEqual([
      {
        sheetName: 'Sheet1',
        cells: [
          { row: 2, column: 0, value: 103 },
          { row: 0, column: 1, value: 1030 },
        ],
      },
    ])
  })

  it('recalculates with the journal applied: a journaled formula is excluded', async () => {
    const { session } = await formulaSession()
    session.setCells({
      sheet: 'Sheet1',
      cells: [
        { ref: 'A1', value: 100 },
        { ref: 'B1', formula: 'A3*10' },
      ],
    })
    await session.save(join(root, 'out.xlsx'))
    // B1's own cache stays empty (the gateway writes formulas without <v>);
    // overlaying an engine value there would freeze the OLD formula's result
    expect(saveCalls[0]?.formulaValues).toEqual([
      { sheetName: 'Sheet1', cells: [{ row: 2, column: 0, value: 103 }] },
    ])
  })

  it('an engine failure degrades to an honest warning, never a failed save', async () => {
    const { session, io } = await formulaSession({
      recalcError: new Error('recalc_busy'),
    })
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 100 }] })
    const saved = await session.save(join(root, 'out.xlsx'))
    expect(io.recalcRequests).toHaveLength(1)
    expect(saveCalls[0]?.formulaValues).toEqual([])
    expect(saved.warnings).toHaveLength(1)
    expect(saved.warnings[0]).toContain('Formula caches were not refreshed')
    expect(saved.warnings[0]).toContain('the formula engine was unavailable')
  })

  it('waits out a lazy formula index and still overlays (BUG-1776 cold flow)', async () => {
    // read_formula_cells returns immediately with whatever the sidecar's lazy
    // background indexer has so far (the sheets app polls it). The canonical
    // open -> edit -> save flow never reads first, so the overlay's single
    // cold call used to always see indexingComplete:false and degrade — the
    // save wrote formulas with no <v> at all. The refresh now polls until the
    // index completes, so a cold save overlays exactly like a warm one.
    const { session, io } = await formulaSession({ formulaIndexingCompleteAfter: 1 })
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 100 }] })
    const saved = await session.save(join(root, 'out.xlsx'))
    expect(io.calls.readFormulaCells.length).toBeGreaterThanOrEqual(2)
    expect(saved.warnings).toEqual([])
    expect(saveCalls[0]?.formulaValues).toEqual([
      {
        sheetName: 'Sheet1',
        cells: [
          { row: 2, column: 0, value: 103 },
          { row: 0, column: 1, value: 1030 },
        ],
      },
    ])
  })

  it('an index that never completes degrades with the reason named', async () => {
    // shrink the wait budget: the first cold reply schedules one 250ms poll,
    // the budget is then exhausted and the save degrades instead of hanging
    setRecalcIndexWaitForTests(10)
    try {
      const { session, io } = await formulaSession({ formulaIndexingComplete: false })
      session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 100 }] })
      const saved = await session.save(join(root, 'out.xlsx'))
      // it retried past the first cold reply before giving up, and the
      // warning states the real effect (no caches, not "stale" caches)
      expect(io.calls.readFormulaCells.length).toBeGreaterThan(1)
      expect(saved.warnings[0]).toContain('Formula caches were not refreshed')
      expect(saved.warnings[0]).toContain('the sheet index did not finish building in time')
      expect(saved.warnings[0]).toContain('see no cached values for the formula cells')
    } finally {
      setRecalcIndexWaitForTests(10_000)
    }
  })

  it('a truncated formula index degrades with the reason named', async () => {
    const { session } = await formulaSession({ formulaTruncated: true })
    session.setCells({ sheet: 'Sheet1', cells: [{ ref: 'A1', value: 100 }] })
    const saved = await session.save(join(root, 'out.xlsx'))
    expect(saved.warnings[0]).toContain('the formula-cell index was truncated or malformed')
  })

  it('a zero-edit save pays no recalculation and stays byte-preserving', async () => {
    const { session, io } = await formulaSession()
    await session.save(join(root, 'out.xlsx'))
    expect(io.calls.recalcCells).toHaveLength(0)
    expect(io.calls.readFormulaCells).toHaveLength(0)
    expect(saveCalls[0]?.formulaValues).toEqual([])
    expect(saveCalls[0]?.edits).toEqual([])
  })
})

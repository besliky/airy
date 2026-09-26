// Integration with the REAL xlsx-sidecar binary (cargo build --release in
// apps/sheets/native/xlsx-engine). Skipped with an explicit reason when the
// binary is not built — the S6 environment has no Rust toolchain, so the
// protocol/session logic is covered by unit tests against a stub, and this
// file guards the true end-to-end path (open -> read -> journal -> gateway
// save -> reopen) plus byte preservation where the binary exists.
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { beforeAll, afterAll, describe, expect, it } from 'vitest'

import { csvToXlsxBuffer, buildWorksheetXml } from '../../../apps/sheets/src/gateway/csv-import.js'
import { buildOdsFixture } from './helpers/odf-fixture.js'
import { patchCentralSizes } from './helpers/zip-bomb.js'
import { findSidecarBinary, sidecarMissingError } from '../src/xlsx/discovery.js'
import { XlsxSession } from '../src/xlsx/session.js'
import { XlsxSidecarClient } from '../src/xlsx/sidecar-client.js'

const binary = await findSidecarBinary()
// Skipped when the binary is missing; the reason rides along in the suite
// title below (building the sidecar needs cargo).
const describeWithBinary = describe.skipIf(!binary)

// A real .xls fixture cannot be generated without a legacy BIFF writer;
// point AIRY_TEST_XLS at one to enable that coverage.
const xlsFixture = process.env.AIRY_TEST_XLS ?? ''
const describeWithXls = describe.skipIf(!binary || xlsFixture === '')

let root: string
let client: XlsxSidecarClient | null = null

beforeAll(async () => {
  if (!binary) return
  root = await mkdtemp(join(tmpdir(), 'airy-xlsx-int-'))
  client = new XlsxSidecarClient(binary)
})

afterAll(async () => {
  client?.stop()
  if (root) await rm(root, { recursive: true, force: true })
})

describeWithBinary(
  'xlsx session against the real sidecar binary ' +
    '(skipped when absent: build with `npm run native:build -w @airy-office/sheets` or set AIRY_XLSX_SIDECAR)',
  () => {
    it('opens, reads values and formulas, journals edits, saves and re-reads', async () => {
      const bookPath = join(root, 'budget.xlsx')
      await writeFile(bookPath, await csvToXlsxBuffer('Region,Sales\nEast,10\nWest,20\n', 'Budget'))

      const session = await XlsxSession.open(bookPath, root, client!)
      const meta = session.meta()
      expect(meta.format).toBe('xlsx')
      expect(meta.sheets.map((sheet) => sheet.name)).toEqual(['Budget'])

      const overview = await session.readWorkbook()
      expect(overview).toMatch(/0\|Budget\|sheet-\d+\|\d+ x \d+/)

      const corner = await session.readWorkbook({ sheet: 'Budget', range: 'A1:B3' })
      expect(corner).toContain('|1|Region|Sales|')
      expect(corner).toContain('|2|East|10|')

      // edit two cells: a constant and a formula over them
      session.setCells({
        sheet: 'Budget',
        cells: [
          { ref: 'B2', value: 15 },
          { ref: 'B4', formula: 'SUM(B2:B3)' },
        ],
      })
      const saved = await session.save(join(root, 'budget-edited.xlsx'))
      expect(saved.path).toBe(join(root, 'budget-edited.xlsx'))
      expect(saved.unchanged).toBe(false)
      expect(saved.touchedEntries.length).toBeGreaterThan(0)

      // reopen the saved copy through a fresh session and verify the writes
      const reopened = await XlsxSession.open(saved.path, root, client!)
      const reread = await reopened.readWorkbook({ sheet: 'Budget', range: 'A1:B4' })
      expect(reread).toContain('|2|East|15|')
      expect(reread).toContain('=SUM(B2:B3)')
      await reopened.close()
      await session.close()
    })

    it('round-trips byte-identically on a zero-edit save (byte preservation)', async () => {
      const bookPath = join(root, 'pristine.xlsx')
      const bytes = await csvToXlsxBuffer('A,B\n1,2\n', 'S')
      await writeFile(bookPath, bytes)
      const session = await XlsxSession.open(bookPath, root, client!)
      const saved = await session.save(join(root, 'pristine-copy.xlsx'))
      expect(saved.unchanged).toBe(true)
      // The save gateway always ensures <calcPr fullCalcOnLoad="1"/> so edited
      // formulas recalculate on open; every OTHER entry raw-copies verbatim
      // (xl/workbook.xml gains the flag when the source lacks it).
      expect(saved.touchedEntries).toEqual(['xl/workbook.xml'])
      const before = await JSZip.loadAsync(bytes)
      const after = await JSZip.loadAsync(await readFile(saved.path))
      for (const name of Object.keys(before.files)) {
        if (before.files[name]!.dir || saved.touchedEntries.includes(name)) continue
        expect(
          Buffer.compare(
            await before.file(name)!.async('nodebuffer'),
            await after.file(name)!.async('nodebuffer'),
          ),
          `untouched entry ${name} changed on a zero-edit save`,
        ).toBe(0)
      }
      await session.close()
    })

    it('refuses a zip bomb declared in the central directory (sidecar budget, SEC-1103)', async () => {
      // A real workbook of a few KiB whose central directory declares
      // gigabytes: the Rust sidecar must refuse it at open — from the
      // declared sizes alone, before decompressing a single entry — and the
      // refusal must reach the session caller unchanged (the Node wrapper
      // adds no generic wrap around sidecar open errors).
      const bombPath = join(root, 'bomb.xlsx')
      const bomb = patchCentralSizes(await csvToXlsxBuffer('A,B\n1,2\n', 'S'), 600 * 1024 * 1024)
      await writeFile(bombPath, bomb)
      await expect(XlsxSession.open(bombPath, root, client!)).rejects.toThrow(
        /Workbook declares \d+ uncompressed bytes across its ZIP entries, .*open budget/,
      )
    })

    it('refuses a too-many-entries workbook naming the count and the budget (BUG-1504 batch)', async () => {
      // the entry-count refusal used to be the vague "Workbook contains too
      // many ZIP entries." — it must name the counter and the budget like the
      // declared-size refusal above (and the docx-engine fence) do
      const zip = await JSZip.loadAsync(await csvToXlsxBuffer('A,B\n1,2\n', 'S'))
      for (let i = 0; i < 10_010; i++) zip.file(`junk/${String(i)}.bin`, 'x')
      const manyPath = join(root, 'many-entries.xlsx')
      await writeFile(manyPath, await zip.generateAsync({ type: 'nodebuffer' }))
      await expect(XlsxSession.open(manyPath, root, client!)).rejects.toThrow(
        /Workbook contains \d+ ZIP entries, above the 10000 entry open budget/,
      )
    })

    // A1=10, A2=20, A3 =SUM(A1:A2), B1 =A3*10; editing A1 to 100 makes the
    // true values 120/1200. The two audit modes: (a) a book with no caches
    // at all — the save used to write formulas without <v>, so openpyxl
    // data_only/pandas read None; (b) a book with stale caches — the save
    // kept A3's cached 30 while the true sum was 120. The package mirrors
    // the app's recalc fixture: IronCalc's importer is strict (complete
    // styles with cellStyles, no whitespace text nodes).
    const buildBook = (cached: boolean): Promise<Buffer> => {
      const zip = new JSZip()
      zip.file(
        '[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          '</Types>',
      )
      zip.file(
        '_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
          '</Relationships>',
      )
      zip.file(
        'xl/workbook.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>' +
          '</workbook>',
      )
      zip.file(
        'xl/_rels/workbook.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
          '</Relationships>',
      )
      zip.file(
        'xl/styles.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
          '</styleSheet>',
      )
      const a3 = cached
        ? '<c r="A3"><f>SUM(A1:A2)</f><v>30</v></c>'
        : '<c r="A3"><f>SUM(A1:A2)</f></c>'
      const b1 = cached ? '<c r="B1"><f>A3*10</f><v>300</v></c>' : '<c r="B1"><f>A3*10</f></c>'
      zip.file(
        'xl/worksheets/sheet1.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<dimension ref="A1:B3"/>' +
          '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
          '<sheetFormatPr defaultRowHeight="15"/>' +
          '<sheetData>' +
          `<row r="1"><c r="A1"><v>10</v></c>${b1}</row>` +
          '<row r="2"><c r="A2"><v>20</v></c></row>' +
          `<row r="3">${a3}</row>` +
          '</sheetData>' +
          '</worksheet>',
      )
      return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    }

    /** the saved sheet XML cell for one ref, for <v> assertions below */
    const savedCellXml = async (path: string, ref: string): Promise<string> => {
      const after = await JSZip.loadAsync(await readFile(path))
      const xml = (await after.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
      return new RegExp(`<c r="${ref}"[^>]*>[\\s\\S]*?</c>`).exec(xml)?.[0] ?? ''
    }

    it('refreshes formula caches on a cold save with no prior read (BUG-1776)', async () => {
      // The canonical agent flow: open -> edit -> save, never reading the
      // sheet. BUG-1776: the overlay's read_formula_cells call raced the
      // sidecar's lazy worksheet index, always saw indexingComplete:false and
      // degraded — the save wrote every formula WITHOUT a <v> (worse than
      // stale: even the file's own caches were dropped). The overlay now
      // waits the index out, so the cold save is as honest as the warm one.
      const bookPath = join(root, 'cold-cache.xlsx')
      await writeFile(bookPath, await buildBook(true))
      const session = await XlsxSession.open(bookPath, root, client!)
      session.setCells({ sheet: 'Data', cells: [{ ref: 'A1', value: 100 }] })
      const saved = await session.save(join(root, 'cold-cache-out.xlsx'))
      expect(saved.unchanged).toBe(false)
      expect(saved.warnings).toEqual([])
      await session.close()
      const staleA3 = await savedCellXml(saved.path, 'A3')
      expect(staleA3).toContain('<f>SUM(A1:A2)</f>')
      expect(staleA3).toContain('<v>120</v>')
      expect(staleA3).not.toContain('<v>30</v>')
      expect(await savedCellXml(saved.path, 'B1')).toContain('<v>1200</v>')
    })

    it('refreshes formula caches on save (BUG-1761: empty and stale <v> modes)', async () => {
      let saveRun = 0
      const savedCell = async (cached: boolean, ref: string): Promise<string> => {
        saveRun += 1
        const bookPath = join(root, `${cached ? 'cache' : 'nocache'}-${String(saveRun)}.xlsx`)
        await writeFile(bookPath, await buildBook(cached))
        const session = await XlsxSession.open(bookPath, root, client!)
        // the warm path: the agent reads before editing (the cold
        // open -> edit -> save flow is covered by the BUG-1776 test above)
        await session.readWorkbook({ sheet: 'Data', range: 'A1:B3' })
        session.setCells({ sheet: 'Data', cells: [{ ref: 'A1', value: 100 }] })
        const saved = await session.save(
          join(root, `${cached ? 'cache' : 'nocache'}-${String(saveRun)}-out.xlsx`),
        )
        expect(saved.unchanged).toBe(false)
        expect(saved.warnings).toEqual([])
        await session.close()
        return savedCellXml(saved.path, ref)
      }

      // (a) cacheless formulas gain a real cached value; the formula stays
      const freshA3 = await savedCell(false, 'A3')
      expect(freshA3).toContain('<f>SUM(A1:A2)</f>')
      expect(freshA3).toContain('<v>120</v>')
      expect(await savedCell(false, 'B1')).toContain('<v>1200</v>')
      // and the edited input itself is a plain constant, not a cache casualty
      expect(await savedCell(false, 'A1')).toContain('<v>100</v>')

      // (b) the stale cache is replaced by the recalculated value
      const staleA3 = await savedCell(true, 'A3')
      expect(staleA3).toContain('<v>120</v>')
      expect(staleA3).not.toContain('<v>30</v>')
    })

    it('refuses an .ods zip bomb declared in the central directory (convert-path fence, SEC-1301)', async () => {
      // Mirror of the xlsx bomb test above for the .ods conversion path:
      // calamine reads .ods through the same ZIP container, so the sidecar
      // runs the same SEC-1103 budgets before convert_workbook decompresses
      // anything. The refusal surfaces inside the session's "Cannot import"
      // wrap with the fence message intact.
      const bombPath = join(root, 'bomb.ods')
      const bomb = patchCentralSizes(await buildOdsFixture(), 600 * 1024 * 1024)
      await writeFile(bombPath, bomb)
      await expect(XlsxSession.open(bombPath, root, client!)).rejects.toThrow(
        /Cannot import .* as \.xlsx: Workbook declares \d+ uncompressed bytes across its ZIP entries, .*open budget/,
      )
    })

    it('editing one sheet keeps every untouched zip entry byte-identical', async () => {
      // a real two-sheet workbook: Main (sheet1.xml) and Data (sheet2.xml)
      const zip = new JSZip()
      zip.file(
        '[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
          '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
          '</Types>',
      )
      zip.file(
        '_rels/.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
          '</Relationships>',
      )
      zip.file(
        'xl/workbook.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<sheets><sheet name="Main" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>',
      )
      zip.file(
        'xl/_rels/workbook.xml.rels',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
          '</Relationships>',
      )
      zip.file('xl/worksheets/sheet1.xml', buildWorksheetXml([['label', 'value']]))
      zip.file(
        'xl/worksheets/sheet2.xml',
        buildWorksheetXml([
          ['alpha', '1'],
          ['beta', '2'],
        ]),
      )
      const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
      const bookPath = join(root, 'multi.xlsx')
      await writeFile(bookPath, bytes)

      const session = await XlsxSession.open(bookPath, root, client!)
      expect(session.meta().sheets.map((sheet) => sheet.name)).toEqual(['Main', 'Data'])
      session.setCells({
        sheet: 'Data',
        cells: [
          { ref: 'B1', value: 99 },
          { ref: 'B2', formula: 'B1*2' },
        ],
      })
      const saved = await session.save(join(root, 'multi-edited.xlsx'))
      expect(saved.unchanged).toBe(false)
      expect(saved.touchedEntries).toContain('xl/worksheets/sheet2.xml')
      expect(saved.touchedEntries).not.toContain('xl/worksheets/sheet1.xml')

      // byte preservation: every file entry the save did not touch is
      // identical (explicit zip directory entries like _rels/ are not
      // raw-copied by the sidecar writer and are skipped)
      const before = await JSZip.loadAsync(bytes)
      const after = await JSZip.loadAsync(await readFile(saved.path))
      for (const name of Object.keys(before.files)) {
        if (before.files[name]!.dir) continue
        const original = before.file(name)
        const result = after.file(name)
        expect(result, `entry ${name} went missing on save`).toBeTruthy()
        if (saved.touchedEntries.includes(name)) continue
        expect(
          Buffer.compare(await original!.async('nodebuffer'), await result!.async('nodebuffer')),
          `untouched entry ${name} changed on save`,
        ).toBe(0)
      }

      // the edits landed on the edited sheet and are readable back
      const reopened = await XlsxSession.open(saved.path, root, client!)
      const reread = await reopened.readWorkbook({ sheet: 'Data', range: 'A1:B2' })
      expect(reread).toContain('|1|alpha|99|')
      expect(reread).toContain('|2|beta|=B1*2|')
      const main = await reopened.readWorkbook({ sheet: 'Main', range: 'A1:B1' })
      expect(main).toContain('|1|label|value|')
      await reopened.close()
      await session.close()
    })

    it('imports .ods via convert_workbook with the styles-lost warning', async () => {
      const odsPath = join(root, 'book.ods')
      const odsBytes = await buildOdsFixture()
      await writeFile(odsPath, odsBytes)
      const session = await XlsxSession.open(odsPath, root, client!)
      const meta = session.meta()
      expect(meta.format).toBe('ods')
      expect(meta.converted).toBe(true)
      expect(meta.warnings[0]).toMatch(/lost on import/)
      const read = await session.readWorkbook({ sheet: 0, range: 'A1:B2' })
      expect(read).toContain('Alpha')
      const saved = await session.save()
      expect(saved.path).toBe(join(root, 'book.xlsx')) // sibling .xlsx
      // the original .ods stays untouched by the default save
      expect(Buffer.compare(await readFile(odsPath), Buffer.from(odsBytes))).toBe(0)
      await session.close()
    })
  },
)

describeWithXls(
  'real legacy .xls import (skipped unless AIRY_TEST_XLS points at a real .xls file — generating a BIFF binary needs a third-party writer)',
  () => {
    it('converts values through convert_workbook', async () => {
      const session = await XlsxSession.open(xlsFixture, root, client!)
      const meta = session.meta()
      expect(meta.format).toBe('xls')
      expect(meta.converted).toBe(true)
      expect(meta.warnings[0]).toMatch(/lost on import/)
      expect(meta.sheets.length).toBeGreaterThan(0)
      await session.close()
    })
  },
)

// Always-run guard so CI output states the skip reasons explicitly.
describe('integration prerequisites', () => {
  it('documents why binary/soffice-dependent suites may be skipped', async () => {
    const missing = binary === null
    if (missing) {
      const message = sidecarMissingError().message
      expect(message).toContain('native:build')
    }
    expect(typeof xlsFixture).toBe('string')
  })
})

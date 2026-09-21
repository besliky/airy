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

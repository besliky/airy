// Integration with the REAL xlsx-sidecar binary (cargo build --release in
// apps/sheets/native/xlsx-engine). Skipped with an explicit reason when the
// binary is not built — the S6 environment has no Rust toolchain, so the
// protocol/session logic is covered by unit tests against a stub, and this
// file guards the true end-to-end path (open -> read -> journal -> gateway
// save -> reopen) plus byte preservation where the binary exists.
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { beforeAll, afterAll, describe, expect, it } from 'vitest'

import { csvToXlsxBuffer } from '../../../apps/sheets/src/gateway/csv-import.js'
import { buildOdsFixture } from './helpers/odf-fixture.js'
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
    '(skipped when absent: build with `npm run native:build -w @genoffice/sheets` or set AIRY_XLSX_SIDECAR)',
  () => {
    it('opens, reads values and formulas, journals edits, saves and re-reads', async () => {
      const bookPath = join(root, 'budget.xlsx')
      await writeFile(bookPath, await csvToXlsxBuffer('Region,Sales\nEast,10\nWest,20\n', 'Budget'))

      const session = await XlsxSession.open(bookPath, root, client!)
      const meta = session.meta()
      expect(meta.format).toBe('xlsx')
      expect(meta.sheets.map((sheet) => sheet.name)).toEqual(['Budget'])

      const overview = await session.readWorkbook()
      expect(overview).toMatch(/0\|Budget\|sheet-0\|\d+ x \d+/)

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
      expect(Buffer.compare(await readFile(saved.path), bytes)).toBe(0)
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

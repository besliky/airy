import { existsSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterAll, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { z } from 'zod'

import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetTableEditRequest } from '../src/gateway/xlsx-gateway'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

/// PAR-INT-55: the table edits (PAR-202) and the native structured-reference
/// resolution (PAR-205) must agree. Rename and Convert to Range rewrite
/// formula text at save time; these checks drive the saved files through the
/// sidecar's recalc channel to prove the rewritten formulas still evaluate,
/// and cover the stripe bake that keeps Convert to Range visually lossless.

const recalcResultSchema = z.object({
  cells: z.array(
    z.object({
      row: z.number(),
      column: z.number(),
      formatted: z.string(),
      number: z.number().optional(),
      isFormula: z.boolean(),
    }),
  ),
  cached: z.boolean(),
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL('../native/xlsx-engine/target/release/' + executable, import.meta.url),
  )
}

const SIDECAR = sidecarBinaryPath()
const HAS_SIDECAR = existsSync(SIDECAR)

// Reference data (hand sums): Sales (A1:B6, totals row 6) Product/Amount
// 10/20/30/40 with a stale totals cache of 100; Prices (C1:C5) with the
// escaped column "Unit Price" 1.5/2.5/3.5/4.5. Formula cells E2:F3 carry
// stale caches so a recalc that keeps them fails the test loudly.
// `sparseStripeRow` drops B4 so the stripe bake has a missing cell to
// create (the sums are only asserted on the complete fixture).
async function buildTableFixture(sparseStripeRow = false): Promise<Buffer> {
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
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
      '<Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
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
  // fills: none, gray125, solid orange; cellXfs: default, orange.
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFED7D31"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
  )
  // A2 carries its own fill (cellXfs 1) — a baked stripe must not take it.
  // B4 is missing — the bake creates it so the stripe stays continuous.
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<dimension ref="A1:F6"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Unit Price</t></is></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr" s="1"><is><t>alpha</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>1.5</v></c><c r="E2"><f>Sales[@Amount]</f><v>99</v></c><c r="F2"><f>SUM(Sales[[#Data],[#Totals],[Amount]])</f><v>99</v></c></row>' +
      '<row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>2.5</v></c><c r="E3"><f>SUM(Sales[Amount])</f><v>999</v></c><c r="F3"><f>SUM(Prices[[Unit Price]])</f><v>999</v></c></row>' +
      '<row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c>' +
      (sparseStripeRow ? '' : '<c r="B4"><v>30</v></c>') +
      '<c r="C4"><v>3.5</v></c><c r="E4"><f>SUM(Sales[#All])</f><v>999</v></c></row>' +
      '<row r="5"><c r="A5" t="inlineStr"><is><t>delta</t></is></c><c r="B5"><v>40</v></c><c r="C5"><v>4.5</v></c></row>' +
      '<row r="6"><c r="A6" t="inlineStr"><is><t>Total</t></is></c><c r="B6"><v>100</v></c></row>' +
      '</sheetData>' +
      '<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>' +
      '</worksheet>',
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/tables/table1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B6" totalsRowCount="1" headerRowCount="1">' +
      '<autoFilter ref="A1:B5"/>' +
      '<tableColumns count="2"><tableColumn id="1" name="Product"/><tableColumn id="2" name="Amount"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
      '</table>',
  )
  zip.file(
    'xl/tables/table2.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Prices" displayName="Prices" ref="C1:C5" headerRowCount="1">' +
      '<autoFilter ref="C1:C5"/>' +
      '<tableColumns count="1"><tableColumn id="1" name="Unit Price"/></tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>' +
      '</table>',
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function runTableEdits(
  tableEdits: SheetTableEditRequest[],
  sparseStripeRow = false,
): Promise<Awaited<ReturnType<typeof applyCellEditsToXlsx>>> {
  return applyCellEditsToXlsx(
    await buildTableFixture(sparseStripeRow),
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    tableEdits,
  )
}

const cleanups: string[] = []

afterAll(async () => {
  await Promise.all(cleanups.map((directory) => rm(directory, { recursive: true, force: true })))
})

async function writeFixture(buffer: Buffer, name: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'xlsx-table-edit-recalc-'))
  cleanups.push(directory)
  const path = join(directory, name)
  await writeFile(path, buffer)
  return path
}

type RecalcEdit = { sheet: string; row: number; column: number; input: string }

async function readNumber(
  client: XlsxSidecarClient,
  path: string,
  edits: readonly RecalcEdit[],
  row: number,
  column: number,
): Promise<number | undefined> {
  const result = recalcResultSchema.parse(
    await client.recalcCells({
      path,
      edits: [...edits],
      reads: [
        {
          sheet: 'Data',
          range: { startRow: row, endRow: row, startColumn: column, endColumn: column },
        },
      ],
    }),
  )
  return result.cells.find((cell) => cell.row === row && cell.column === column)?.number
}

describe('table rename keeps structured references computing', { skip: !HAS_SIDECAR }, () => {
  it('rewrites bare, @, and special-item forms and the engine still resolves them', async () => {
    const mutation = await runTableEdits([
      { sheetName: 'Data', tableName: 'Sales', rename: 'Revenue' },
    ])
    const worksheet = await JSZip.loadAsync(mutation.buffer).then((archive) =>
      archive.file('xl/worksheets/sheet1.xml')!.async('text'),
    )
    // The stored formulas keep the user-facing @ text under the new name.
    expect(worksheet).toContain('<f>Revenue[@Amount]</f>')
    expect(worksheet).toContain('<f>SUM(Revenue[Amount])</f>')
    expect(worksheet).toContain('<f>SUM(Revenue[#All])</f>')
    // Multi-specifier groups swap the name too (the file text stays valid
    // for Excel even though the 0.8.3 engine cannot lex this form yet).
    expect(worksheet).toContain('<f>SUM(Revenue[[#Data],[#Totals],[Amount]])</f>')
    // Other tables and their references are untouched.
    expect(worksheet).toContain('<f>SUM(Prices[[Unit Price]])</f>')
    const table = await JSZip.loadAsync(mutation.buffer).then((archive) =>
      archive.file('xl/tables/table1.xml')!.async('text'),
    )
    expect(table).toContain('displayName="Revenue"')

    // Engine-level: the renamed references evaluate (stale caches were 99/999).
    const path = await writeFixture(mutation.buffer, 'renamed.xlsx')
    const client = new XlsxSidecarClient(SIDECAR)
    try {
      expect(await readNumber(client, path, [], 1, 4)).toBe(10) // Revenue[@Amount]
      expect(await readNumber(client, path, [], 2, 4)).toBe(100) // SUM(Revenue[Amount])
      expect(await readNumber(client, path, [], 3, 4)).toBe(200) // SUM(Revenue[#All])
      expect(await readNumber(client, path, [], 2, 5)).toBe(12) // Prices untouched
    } finally {
      client.stop()
    }
  })
})

describe('convert to range makes structured references computable', { skip: !HAS_SIDECAR }, () => {
  it('rewrites them to A1 and drops only the converted table', async () => {
    const mutation = await runTableEdits([
      { sheetName: 'Data', tableName: 'Sales', convertToRange: true },
    ])
    const worksheet = await JSZip.loadAsync(mutation.buffer).then((archive) =>
      archive.file('xl/worksheets/sheet1.xml')!.async('text'),
    )
    expect(worksheet).toContain('<f>B2</f>')
    expect(worksheet).toContain('<f>SUM(B2:B5)</f>')
    expect(worksheet).toContain('<f>SUM(B2:B6)</f>') // [[#Data],[#Totals]] union
    // Prices survives with its references untouched.
    expect(worksheet).toContain('<f>SUM(Prices[[Unit Price]])</f>')
    expect(worksheet).toContain('<tableParts count="1">')

    // Engine-level: the A1 formulas evaluate, including other-table refs.
    const path = await writeFixture(mutation.buffer, 'converted.xlsx')
    const client = new XlsxSidecarClient(SIDECAR)
    try {
      expect(await readNumber(client, path, [], 1, 4)).toBe(10) // B2
      expect(await readNumber(client, path, [], 2, 4)).toBe(100) // SUM(B2:B5)
      expect(await readNumber(client, path, [], 1, 5)).toBe(200) // SUM(B2:B6)
      expect(await readNumber(client, path, [], 2, 5)).toBe(12) // SUM(Prices[[Unit Price]])
    } finally {
      client.stop()
    }
  })
})

describe('rename → convert in one session (BUG-1750)', { skip: !HAS_SIDECAR }, () => {
  it('converts the renamed references to A1; nothing points at the deleted table', async () => {
    // The audited chain: Rename (Apply) then Convert to Range (Apply) in one
    // session — the journal merges both into a single edit keyed by the
    // original name. The conversion must match the post-rename token, or the
    // structured references ride to the file untouched while their table part
    // is deleted (#NAME? across the workbook in Excel).
    const mutation = await runTableEdits([
      { sheetName: 'Data', tableName: 'Sales', rename: 'Revenue', convertToRange: true },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(zip.file('xl/tables/table1.xml')).toBeNull()
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    // Every reference became its A1 equivalent, under either token.
    expect(worksheet).toContain('<f>B2</f>') // @-shorthand
    expect(worksheet).toContain('<f>SUM(B2:B5)</f>') // data body
    expect(worksheet).toContain('<f>SUM(A1:B6)</f>') // #All (whole table)
    expect(worksheet).toContain('<f>SUM(B2:B6)</f>') // [[#Data],[#Totals]] union
    expect(worksheet).not.toContain('Sales[')
    expect(worksheet).not.toContain('Revenue')
    // Prices survives with its references untouched.
    expect(worksheet).toContain('<f>SUM(Prices[[Unit Price]])</f>')
    expect(worksheet).toContain('<tableParts count="1">')

    // Engine-level: the A1 formulas evaluate (the stale caches were 99/999).
    const path = await writeFixture(mutation.buffer, 'renamed-converted.xlsx')
    const client = new XlsxSidecarClient(SIDECAR)
    try {
      expect(await readNumber(client, path, [], 1, 4)).toBe(10) // B2
      expect(await readNumber(client, path, [], 2, 4)).toBe(100) // SUM(B2:B5)
      expect(await readNumber(client, path, [], 3, 4)).toBe(200) // SUM(B1:B6)
      expect(await readNumber(client, path, [], 1, 5)).toBe(200) // SUM(B2:B6)
      expect(await readNumber(client, path, [], 2, 5)).toBe(12) // Prices untouched
    } finally {
      client.stop()
    }
  })
})

describe('convert to range bakes the row stripes', () => {
  it('writes the stripe fill into alternating body cells and creates missing cells', async () => {
    const mutation = await runTableEdits(
      [
        {
          sheetName: 'Data',
          tableName: 'Sales',
          convertToRange: true,
          stripeFill: '#B8CCE4',
        },
      ],
      true,
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const styles = await zip.file('xl/styles.xml')!.async('text')
    // One new solid fill + one new xf derived from the default format.
    expect(styles).toContain('<fills count="4">')
    expect(styles).toContain('<fgColor rgb="FFB8CCE4"/>')
    expect(styles).toContain(
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>',
    )
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    // Stripe rows 2 and 4 (physical parity from the first data row).
    expect(worksheet).toContain('<c r="B2" s="2"><v>10</v></c>')
    expect(worksheet).toContain('<c r="A4" t="inlineStr" s="2"><is><t>gamma</t></is></c>')
    // The missing cell was created so the stripe stays continuous.
    expect(worksheet).toContain('<c r="B4" s="2"/>')
    // A cell with its own fill keeps it; non-stripe rows are untouched.
    expect(worksheet).toContain('<c r="A2" t="inlineStr" s="1"><is><t>alpha</t></is></c>')
    expect(worksheet).toContain('<c r="B3"><v>20</v></c>')
    expect(worksheet).toContain('<c r="B6"><v>100</v></c>')
  })

  it('leaves cell formatting untouched without a stripe fill', async () => {
    const mutation = await runTableEdits([
      { sheetName: 'Data', tableName: 'Sales', convertToRange: true },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const styles = await zip.file('xl/styles.xml')!.async('text')
    expect(styles).toContain('<fills count="3">')
    expect(styles).not.toContain('B8CCE4')
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(worksheet).toContain('<c r="B2"><v>10</v></c>')
    expect(worksheet).toContain('<c r="B4"><v>30</v></c>')
  })
})

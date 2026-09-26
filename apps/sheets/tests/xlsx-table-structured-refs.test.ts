import { existsSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterAll, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { z } from 'zod'

import { applyCellEditsToXlsx, type CellEdit } from '../src/gateway/xlsx-gateway'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

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

/// Structured references (PAR-205): formulas like `=SUM(Sales[Amount])` and
/// Excel's `@` this-row shorthand must (a) evaluate through the sidecar's
/// IronCalc recalc channel against the workbook's table parts, and (b)
/// round-trip byte-verbatim through the save path — stored formulas always
/// keep the user-facing `@` text, only the compute model sees the rewrite.

const PYTHON = ['/usr/bin/python3', '/usr/local/bin/python3'].find((path) => existsSync(path))
const HAS_OPENPYXL =
  PYTHON !== undefined &&
  (() => {
    try {
      execFileSync(PYTHON, ['-c', 'import openpyxl'], { stdio: 'ignore' })
      return true
    } catch {
      return false
    }
  })()

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL('../native/xlsx-engine/target/release/' + executable, import.meta.url),
  )
}

const SIDECAR = sidecarBinaryPath()
const HAS_SIDECAR = existsSync(SIDECAR)

// Reference data: Sales (A1:B6, totals row 6) Product/Amount with amounts
// 10/20/30/40 and a totals row cached at 100; Prices (C1:C5) with the
// escaped column name "Unit Price" (1.5/2.5/3.5/4.5). Hand sums:
// Sales[Amount]=100, Prices[[Unit Price]]=12, #All bands add the totals row.
async function buildTableFixture(): Promise<Buffer> {
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
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<dimension ref="A1:C6"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Unit Price</t></is></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>1.5</v></c></row>' +
      '<row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>2.5</v></c></row>' +
      '<row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c><c r="B4"><v>30</v></c><c r="C4"><v>3.5</v></c></row>' +
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

const cleanups: string[] = []

afterAll(async () => {
  await Promise.all(cleanups.map((directory) => rm(directory, { recursive: true, force: true })))
})

async function writeFixture(name: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'xlsx-structured-refs-'))
  cleanups.push(directory)
  const path = join(directory, name)
  await writeFile(path, await buildTableFixture())
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

describe('structured references through the recalc channel', { skip: !HAS_SIDECAR }, () => {
  it('computes the reference table sums from user edits', async () => {
    const path = await writeFixture('sales.xlsx')
    const client = new XlsxSidecarClient(SIDECAR)
    try {
      // Sales[Amount] = 10+20+30+40; the totals row is excluded.
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 7, column: 4, input: '=SUM(Sales[Amount])' }],
          7,
          4,
        ),
      ).toBe(100)
      // Escaped column with spaces, whole-column sum: 1.5+2.5+3.5+4.5.
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 7, column: 5, input: '=SUM(Prices[[Unit Price]])' }],
          7,
          5,
        ),
      ).toBe(12)
    } finally {
      client.stop()
    }
  })

  it('resolves the @ this-row shorthand and special items from edits', async () => {
    const path = await writeFixture('sales.xlsx')
    const client = new XlsxSidecarClient(SIDECAR)
    try {
      // Sheet row 2 (wire row 1): the @ shorthand picks that row.
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 1, column: 4, input: '=Sales[@Amount]' }],
          1,
          4,
        ),
      ).toBe(10)
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 1, column: 5, input: '=Prices[@[Unit Price]]' }],
          1,
          5,
        ),
      ).toBe(1.5)
      // A this-row column range: B2:B2 = 10.
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 1, column: 6, input: '=SUM(Sales[@[Amount]:[Amount]])' }],
          1,
          6,
        ),
      ).toBe(10)
      // Special items: #All spans headers + data + totals (SUM skips texts),
      // #Headers counts the two headers, #Data the eight data cells, and
      // #Totals picks the totals row alone.
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 3, column: 4, input: '=SUM(Sales[[#All],[Amount]])' }],
          3,
          4,
        ),
      ).toBe(200)
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 4, column: 4, input: '=COUNTA(Sales[#Headers])' }],
          4,
          4,
        ),
      ).toBe(2)
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 5, column: 4, input: '=COUNTA(Sales[#Data])' }],
          5,
          4,
        ),
      ).toBe(8)
      expect(
        await readNumber(
          client,
          path,
          [{ sheet: 'Data', row: 6, column: 4, input: '=SUM(Sales[[#Totals],[Amount]])' }],
          6,
          4,
        ),
      ).toBe(100)
    } finally {
      client.stop()
    }
  })

  it('computes file formulas that use the @ shorthand on a cold open', async () => {
    // The fixture formula cell carries a stale cache of 99: without the
    // this-row rewrite the cold import would keep it at the stale value.
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-structured-refs-'))
    cleanups.push(directory)
    const path = join(directory, 'at-file.xlsx')
    const zip = await JSZip.loadAsync(await buildTableFixture())
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    zip.file(
      'xl/worksheets/sheet1.xml',
      sheet.replace(
        '<c r="B2"><v>10</v></c>',
        '<c r="B2"><v>10</v></c><c r="E2"><f>Sales[@Amount]</f><v>99</v></c>',
      ),
    )
    await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
    const client = new XlsxSidecarClient(SIDECAR)
    try {
      expect(await readNumber(client, path, [], 1, 4)).toBe(10)
    } finally {
      client.stop()
    }
  })
})

function valueEdit(row: number, column: number, value: string): CellEdit {
  return { sheetName: 'Data', row, column, writeValue: true, cell: { value } }
}

describe('structured reference round-trip through the save path', () => {
  it('keeps table formulas and table parts byte-verbatim on save', async () => {
    const mutation = await applyCellEditsToXlsx(await buildTableFixture(), [
      valueEdit(0, 4, 'untouched'),
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(worksheet).toContain(
      '<tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts>',
    )
    const table = await zip.file('xl/tables/table1.xml')!.async('text')
    expect(table).toContain('displayName="Sales"')
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
    // The table parts and worksheet body are not touched by an unrelated
    // value edit: an Excel file with structured references survives a save.
    expect(after.get('xl/tables/table1.xml')).toBe(before.get('xl/tables/table1.xml'))
    expect(after.get('xl/tables/table2.xml')).toBe(before.get('xl/tables/table2.xml'))
  })

  it('writes a structured reference formula entered by the user as-is', async () => {
    const mutation = await applyCellEditsToXlsx(await buildTableFixture(), [
      {
        sheetName: 'Data',
        row: 7,
        column: 4,
        writeValue: true,
        cell: { value: null, formula: '=SUM(Sales[@[Amount]:[Amount]])' },
      },
    ])
    const worksheet = await JSZip.loadAsync(mutation.buffer).then((archive) =>
      archive.file('xl/worksheets/sheet1.xml')!.async('text'),
    )
    expect(worksheet).toContain('<f>SUM(Sales[@[Amount]:[Amount]])</f>')
  })

  it.skipIf(!HAS_OPENPYXL)('openpyxl loads the saved workbook with its tables', async () => {
    const mutation = await applyCellEditsToXlsx(await buildTableFixture(), [
      valueEdit(0, 4, 'untouched'),
    ])
    const directory = await mkdtemp(join(tmpdir(), 'airy-structured-refs-'))
    try {
      const path = join(directory, 'structured.xlsx')
      await writeFile(path, mutation.buffer)
      // Loading is the verification: openpyxl parses the table parts, the
      // rels graph and the @ formulas like an external consumer would.
      const output = execFileSync(
        PYTHON!,
        [
          '-c',
          [
            'import openpyxl, json, sys',
            'wb = openpyxl.load_workbook(sys.argv[1])',
            'ws = wb["Data"]',
            'tables = {name: str(ref) for name, ref in ws.tables.items()}',
            'print(json.dumps({"tables": tables, "b2": ws["B2"].value}))',
          ].join('\n'),
          path,
        ],
        { stdio: 'pipe' },
      ).toString()
      const parsed = JSON.parse(output) as { tables: Record<string, string>; b2: number }
      expect(parsed.tables.Sales).toBe('A1:B6')
      expect(parsed.tables.Prices).toBe('C1:C5')
      expect(parsed.b2).toBe(10)
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })
})

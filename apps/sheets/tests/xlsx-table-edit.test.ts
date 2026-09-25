import { existsSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { execFileSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '../src/gateway/xlsx-gateway'
import type { SheetTableAddition, SheetTableEditRequest } from '../src/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

/// Gateway coverage for editing Excel tables stored in the file (PAR-202):
/// Resize, Rename (with structured-reference rewrites), style, and Convert
/// to Range, plus an openpyxl round-trip of the saved package.

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

const TABLE_USERS: SheetTableAddition = {
  sheetName: 'Data',
  area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
  name: 'Users',
  columnNames: ['Country', 'Users'],
  bandedRows: true,
}

const FORMULA_CELL = '<row r="2"><c r="D2"><f>SUM(Users[Users])</f></c></row>'
const DEFINED_NAME = '<definedName name="MyRef">SUM(Users[Country])</definedName>'

/// The edit fixture with the Users table added and a structured-reference
/// formula cell plus a table-referencing defined name patched in, so the
/// rename/convert rewrites have real targets.
async function tableSource(withFormula = true): Promise<Buffer> {
  const fixture = await buildEditFixture()
  const source = await createBufferEntrySource(fixture)
  const plan = await planCellEditsToXlsx(
    source,
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
    [TABLE_USERS],
  )
  const buffer = (await assembleWithJsZip(fixture, plan)).buffer
  if (!withFormula) return buffer
  const zip = await JSZip.loadAsync(buffer)
  const worksheet = (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(
    '<row r="3"',
    `${FORMULA_CELL}<row r="3"`,
  )
  zip.file('xl/worksheets/sheet1.xml', worksheet)
  const workbook = (await zip.file('xl/workbook.xml')!.async('string')).replace(
    '</definedNames>',
    `${DEFINED_NAME}</definedNames>`,
  )
  zip.file('xl/workbook.xml', workbook)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function planWith(
  source: Buffer,
  tableEdits: SheetTableEditRequest[],
  structuralOps: Parameters<typeof planCellEditsToXlsx>[2] = [],
) {
  const entrySource = await createBufferEntrySource(source)
  return planCellEditsToXlsx(
    entrySource,
    [],
    structuralOps,
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
    [],
    [],
    [],
    [],
    [],
    [],
    [],
    null,
    null,
    [],
    [],
    [],
    tableEdits,
  )
}

describe('table resize', () => {
  it('grows rows and columns with an Excel-style name for the new column', async () => {
    const plan = await planWith(await tableSource(), [
      {
        sheetName: 'Data',
        tableName: 'Users',
        resize: { area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 2 } },
      },
    ])
    const tableXml = plan.replaced.get('xl/tables/table1.xml')!
    expect(tableXml).toContain('ref="A1:C6"')
    expect(tableXml).toContain('<autoFilter ref="A1:C6"/>')
    expect(tableXml).toContain('<tableColumns count="3">')
    expect(tableXml).toContain('<tableColumn id="3" name="Column3"/>')
    // The header row stays anchored; shrinking never touches cells.
    expect(plan.removedEntries).toHaveLength(0)
  })

  it('shrinks rows and columns, dropping only trailing tableColumns', async () => {
    const plan = await planWith(await tableSource(), [
      {
        sheetName: 'Data',
        tableName: 'Users',
        resize: { area: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 } },
      },
    ])
    const tableXml = plan.replaced.get('xl/tables/table1.xml')!
    expect(tableXml).toContain('ref="A1:A2"')
    expect(tableXml).toContain('<tableColumns count="1">')
    expect(tableXml).toContain('<tableColumn id="1" name="Country"/>')
    expect(tableXml).not.toContain('<tableColumn id="2"')
  })

  it('refuses to move the header cell or drop the last data row', async () => {
    await expect(
      planWith(await tableSource(), [
        {
          sheetName: 'Data',
          tableName: 'Users',
          resize: { area: { startRow: 1, startColumn: 0, endRow: 5, endColumn: 1 } },
        },
      ]),
    ).rejects.toThrow(/top-left corner cannot move/)
    await expect(
      planWith(await tableSource(), [
        {
          sheetName: 'Data',
          tableName: 'Users',
          resize: { area: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 } },
        },
      ]),
    ).rejects.toThrow(/at least one data row/)
  })

  it('refuses to grow into another table', async () => {
    const fixture = await buildEditFixture()
    const source = await createBufferEntrySource(fixture)
    const plan = await planCellEditsToXlsx(
      source,
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
      [
        TABLE_USERS,
        {
          sheetName: 'Data',
          area: { startRow: 0, startColumn: 4, endRow: 3, endColumn: 5 },
          name: 'Costs',
          columnNames: ['Item', 'Cost'],
          bandedRows: true,
        },
      ],
    )
    const grown = (await assembleWithJsZip(fixture, plan)).buffer
    await expect(
      planWith(grown, [
        {
          sheetName: 'Data',
          tableName: 'Users',
          resize: { area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 5 } },
        },
      ]),
    ).rejects.toThrow(/would overlap an existing table/)
  })
})

describe('table rename', () => {
  it('renames the part and rewrites structured references in formulas and names', async () => {
    const plan = await planWith(await tableSource(), [
      { sheetName: 'Data', tableName: 'Users', rename: 'Clients' },
    ])
    const tableXml = plan.replaced.get('xl/tables/table1.xml')!
    expect(tableXml).toContain('name="Clients" displayName="Clients"')
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(worksheet).toContain('<f>SUM(Clients[Users])</f>')
    expect(worksheet).not.toContain('SUM(Users')
    const workbook = plan.replaced.get('xl/workbook.xml')!
    expect(workbook).toContain('SUM(Clients[Country])')
  })

  it('enforces Excel name rules and uniqueness', async () => {
    const source = await tableSource()
    await expect(
      planWith(source, [{ sheetName: 'Data', tableName: 'Users', rename: 'My Table' }]),
    ).rejects.toThrow(/no spaces/)
    await expect(
      planWith(source, [{ sheetName: 'Data', tableName: 'Users', rename: 'A1' }]),
    ).rejects.toThrow(/cell reference/)
    await expect(
      planWith(source, [{ sheetName: 'Data', tableName: 'Users', rename: 'Total' }]),
    ).rejects.toThrow(/collides with a defined name/)
  })
})

describe('table style edit', () => {
  it('rewrites the style name and banding flag', async () => {
    const plan = await planWith(await tableSource(), [
      {
        sheetName: 'Data',
        tableName: 'Users',
        style: { style: 'TableStyleLight9', bandedRows: false },
      },
    ])
    const tableXml = plan.replaced.get('xl/tables/table1.xml')!
    expect(tableXml).toContain('name="TableStyleLight9"')
    expect(tableXml).toContain('showRowStripes="0"')
  })
})

describe('convert to range', () => {
  it('removes the table and converts structured references to A1', async () => {
    const buffer = await tableSource()
    const plan = await planWith(buffer, [
      { sheetName: 'Data', tableName: 'Users', convertToRange: true },
    ])
    // The part, its relationship, its worksheet hookup, and its content type
    // are all gone.
    expect(plan.removedEntries).toContain('xl/tables/table1.xml')
    expect(plan.added.has('xl/tables/table1.xml')).toBe(false)
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(worksheet).not.toContain('<tablePart')
    expect(worksheet).not.toContain('<tableParts')
    const rels = plan.replaced.get('xl/worksheets/_rels/sheet1.xml.rels')!
    expect(rels).not.toContain('tables/table1.xml')
    const contentTypes = plan.replaced.get('[Content_Types].xml')!
    expect(contentTypes).not.toContain('/xl/tables/table1.xml')
    // Structured references became A1 equivalents; the data survives.
    expect(worksheet).toContain('<f>SUM(B2:B4)</f>')
    expect(worksheet).toContain('t="s"')
    const workbook = plan.replaced.get('xl/workbook.xml')!
    // Workbook-level names take the table's sheet qualification.
    expect(workbook).toContain('SUM(Data!A2:A4)')
  })

  it('fails closed when the table does not exist', async () => {
    await expect(
      planWith(await tableSource(), [
        { sheetName: 'Data', tableName: 'Ghost', convertToRange: true },
      ]),
    ).rejects.toThrow(/was not found/)
  })
})

describe('table edit guards', () => {
  it('refuses to ride with row/column shifts on the same sheet', async () => {
    await expect(
      planWith(
        await tableSource(),
        [{ sheetName: 'Data', tableName: 'Users', rename: 'Clients' }],
        [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 8, count: 1 }] }],
      ),
    ).rejects.toThrow(/save the table first/)
  })
})

describe('round-trip', () => {
  it.skipIf(!HAS_OPENPYXL)('openpyxl sees the resized, renamed table', async () => {
    const buffer = await tableSource(false)
    const plan = await planWith(buffer, [
      {
        sheetName: 'Data',
        tableName: 'Users',
        rename: 'Clients',
        resize: { area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 2 } },
      },
    ])
    const saved = (await assembleWithJsZip(buffer, plan)).buffer
    const dir = await mkdtemp(join(tmpdir(), 'airy-table-edit-'))
    try {
      const path = join(dir, 'edited.xlsx')
      await writeFile(path, saved)
      const out = execFileSync(
        PYTHON!,
        [
          '-c',
          [
            'import openpyxl, sys',
            'wb = openpyxl.load_workbook(sys.argv[1])',
            'ws = wb["Data"]',
            'assert list(ws.tables) == ["Clients"], ws.tables',
            'table = ws.tables["Clients"]',
            'assert table.ref == "A1:C6", table.ref',
            'assert [c.name for c in table.tableColumns] == ["Country", "Users", "Column3"]',
            'assert ws["A1"].value == "Hello"',
            'print("ok")',
          ].join('; '),
          path,
        ],
        { stdio: 'pipe' },
      )
      expect(out.toString()).toContain('ok')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it.skipIf(!HAS_OPENPYXL)('openpyxl sees plain data after Convert to Range', async () => {
    const buffer = await tableSource()
    const plan = await planWith(buffer, [
      { sheetName: 'Data', tableName: 'Users', convertToRange: true },
    ])
    const saved = (await assembleWithJsZip(buffer, plan)).buffer
    // Structural re-check of the saved bytes: no table part remains.
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('xl/tables/table1.xml')).toBeNull()
    const dir = await mkdtemp(join(tmpdir(), 'airy-table-convert-'))
    try {
      const path = join(dir, 'converted.xlsx')
      await writeFile(path, saved)
      const out = execFileSync(
        PYTHON!,
        [
          '-c',
          [
            'import openpyxl, sys',
            'wb = openpyxl.load_workbook(sys.argv[1])',
            'ws = wb["Data"]',
            'assert len(ws.tables) == 0, ws.tables',
            'assert ws["A1"].value == "Hello"',
            'assert ws["D2"].value == "=SUM(B2:B4)"',
            'print("ok")',
          ].join('; '),
          path,
        ],
        { stdio: 'pipe' },
      )
      expect(out.toString()).toContain('ok')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })
})

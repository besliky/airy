import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  planCellEditsToXlsx,
  createBufferEntrySource,
  assembleWithJsZip,
} from '../src/gateway/xlsx-gateway'
import type { SheetFilterState } from '../src/gateway/xlsx-filter'
import type { SheetSlicerAddition, SheetTableAddition } from '../src/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'

function tableAddition(): SheetTableAddition {
  return {
    sheetName: 'Data',
    area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
    name: 'Users',
    columnNames: ['Country', 'Users'],
    bandedRows: true,
  }
}

function slicerAddition(overrides: Partial<SheetSlicerAddition> = {}): SheetSlicerAddition {
  return { sheetName: 'Data', tableName: 'Users', colId: 0, ...overrides }
}

async function planWith(
  tables: SheetTableAddition[],
  slicers: SheetSlicerAddition[] = [],
  buffer?: Buffer,
  filterStates: SheetFilterState[] = [],
) {
  const source = await createBufferEntrySource(buffer ?? (await buildEditFixture()))
  return planCellEditsToXlsx(
    source,
    [],
    [],
    [],
    undefined,
    filterStates,
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    tables,
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
    [],
    slicers,
  )
}

async function unzip(buffer: Buffer): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(buffer)
  const entries = new Map<string, string>()
  for (const name of Object.keys(zip.files)) {
    const file = zip.file(name)
    if (file) entries.set(name, await file.async('string'))
  }
  return entries
}

describe('table slicer persistence', () => {
  it('writes the slicer part, slicerCache with tableSlicerCache, and every workbook/worksheet hook', async () => {
    const plan = await planWith([tableAddition()], [slicerAddition()])

    const slicerXml = plan.added.get('xl/slicers/slicer1.xml')
    expect(slicerXml).toBeDefined()
    expect(slicerXml).toContain('<slicer name="Country" cache="Slicer_Country" caption="Country"')
    expect(slicerXml).toContain('rowHeight=')
    expect(slicerXml).toContain('http://schemas.microsoft.com/office/spreadsheetml/2009/9/main')

    const cacheXml = plan.added.get('xl/slicerCaches/slicerCache1.xml')
    expect(cacheXml).toBeDefined()
    expect(cacheXml).toContain('name="Slicer_Country" sourceName="Country"')
    expect(cacheXml).toContain(
      `ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${X15_NS}"`,
    )
    expect(cacheXml).toContain('<x15:tableSlicerCache tableId="1" column="1"/>')

    const workbookXml = plan.replaced.get('xl/workbook.xml')
    expect(workbookXml).toContain(
      `ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${X15_NS}"`,
    )
    expect(workbookXml).toContain('<x15:slicerCache r:id=')
    expect(workbookXml).toContain('<definedName name="Slicer_Country">#N/A</definedName>')

    const workbookRels = plan.replaced.get('xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain(
      'Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache"',
    )
    expect(workbookRels).toContain('Target="slicerCaches/slicerCache1.xml"')

    const worksheetXml = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheetXml).toContain(
      `ext uri="{3A4CF648-6AED-4f9d-8E56-E3E7D4C4D3F7}" xmlns:x15="${X15_NS}"`,
    )
    expect(worksheetXml).toContain('<x15:slicerList><x15:slicer r:id=')

    const sheetRels = plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain(
      'Type="http://schemas.microsoft.com/office/2007/relationships/slicer"',
    )
    expect(sheetRels).toContain('Target="../slicers/slicer1.xml"')

    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain(
      'PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"',
    )
    expect(contentTypes).toContain(
      'PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"',
    )
  })

  it('skips an identical binding and allocates fresh names for a second column', async () => {
    const plan = await planWith(
      [tableAddition()],
      [slicerAddition(), slicerAddition(), slicerAddition({ colId: 1 })],
    )
    expect(plan.added.has('xl/slicers/slicer1.xml')).toBe(true)
    // The duplicate (table, column) slicer is deduplicated; the second column
    // gets its own parts and a distinct cache name.
    expect(plan.added.has('xl/slicers/slicer2.xml')).toBe(true)
    const slicer2 = plan.added.get('xl/slicers/slicer2.xml')
    expect(slicer2).toContain('cache="Slicer_Users"')
    const cache2 = plan.added.get('xl/slicerCaches/slicerCache2.xml')
    expect(cache2).toContain('<x15:tableSlicerCache tableId="1" column="2"/>')
  })

  it('fails closed when the column is outside the table or the table is unknown', async () => {
    await expect(planWith([tableAddition()], [slicerAddition({ colId: 5 })])).rejects.toThrow(
      /outside table/,
    )
    await expect(
      planWith([tableAddition()], [slicerAddition({ tableName: 'Missing' })]),
    ).rejects.toThrow(/was not found/)
  })

  it('round-trips: filter → save keeps the slicer parts byte-present', async () => {
    const first = await planWith([tableAddition()], [slicerAddition()])
    const saved = (await assembleWithJsZip(await buildEditFixture(), first)).buffer

    // Reopen the saved file and save again with a table-owned filter snapshot
    // (the panel rewrite path): slicer parts must survive untouched.
    const filterStates: SheetFilterState[] = [
      {
        sheetName: 'Data',
        filter: {
          range: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
          columns: [{ colId: 0, values: ['de'] }],
        },
        hiddenRows: [1, 3],
        visibilityRange: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
        tableName: 'Users',
      },
    ]
    const second = await planWith([], [], saved, filterStates)
    const entries = await unzip((await assembleWithJsZip(saved, second)).buffer)

    expect(entries.get('xl/slicers/slicer1.xml')).toContain('cache="Slicer_Country"')
    expect(entries.get('xl/slicerCaches/slicerCache1.xml')).toContain(
      '<x15:tableSlicerCache tableId="1" column="1"/>',
    )
    expect(entries.get('xl/workbook.xml')).toContain('<x15:slicerCache r:id=')
    expect(entries.get('xl/worksheets/sheet1.xml')).toContain('<x15:slicerList>')

    // The criteria went into the TABLE part's autoFilter — not the worksheet.
    const tableXml = entries.get('xl/tables/table1.xml')
    expect(tableXml).toContain(
      '<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="de"/></filters></filterColumn></autoFilter>',
    )
    const worksheetXml = entries.get('xl/worksheets/sheet1.xml')
    expect(worksheetXml).not.toMatch(/<autoFilter\b[^>]*>\s*<filterColumn/)
    // hiddenRows [1, 3] (0-based) → sheet rows 2 and 4 carry hidden="1".
    expect(worksheetXml).toContain('<row r="2" hidden="1"/>')
    expect(worksheetXml).toContain('<row r="4" hidden="1"/>')
  })

  it('removes the table autoFilter when the table-owned filter is cleared', async () => {
    const first = await planWith([tableAddition()], [])
    const saved = (await assembleWithJsZip(await buildEditFixture(), first)).buffer
    const filterStates: SheetFilterState[] = [
      {
        sheetName: 'Data',
        filter: null,
        hiddenRows: [],
        visibilityRange: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
        tableName: 'Users',
      },
    ]
    const second = await planWith([], [], saved, filterStates)
    const entries = await unzip((await assembleWithJsZip(saved, second)).buffer)
    expect(entries.get('xl/tables/table1.xml')).not.toContain('<autoFilter')
    // Worksheet-level rows are unhidden inside the visibility span.
    expect(entries.get('xl/worksheets/sheet1.xml')).not.toContain('hidden="1"')
  })
})

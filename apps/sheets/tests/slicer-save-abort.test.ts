/**
 * BUG-1751 regression: table slicers and their table-owned filter criteria
 * must not abort the save when the table is renamed or converted to a range
 * in the same session. The renderer resolves the slicer's table name through
 * the pending tableEdits before the request (see table-slicer.test.ts); the
 * gateway must accept the resolved shapes: slicer additions and the deferred
 * table-filter write run after the rename, and a Convert to Range save with
 * a cleared criteria snapshot keeps the filter's hidden rows while the table
 * part (and its slicers) go away.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { createBufferEntrySource, planCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetFilterState } from '../src/gateway/xlsx-filter'
import type { SheetSlicerAddition, SheetTableEditRequest } from '../src/gateway/xlsx-gateway'

const TABLE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" ' +
  'name="Sales" displayName="Sales" ref="A1:B4">' +
  '<autoFilter ref="A1:B4"/>' +
  '<tableColumns count="2"><tableColumn id="1" name="Product"/><tableColumn id="2" name="Amount"/>' +
  '</tableColumns></table>'

const WORKSHEET_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheetData>' +
  '<row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>' +
  '<row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>1</v></c></row>' +
  '<row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3"><v>2</v></c></row>' +
  '<row r="4"><c r="A4" t="inlineStr"><is><t>alpha</t></is></c><c r="B4"><v>3</v></c></row>' +
  '</sheetData>' +
  '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' +
  '</worksheet>'

async function salesWorkbook(): Promise<Buffer> {
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
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
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
      '<fonts count="1"><font><sz val="11"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
      '<cellXfs count="1"><xf/></cellXfs>' +
      '</styleSheet>',
  )
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET_XML)
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/tables/table1.xml', TABLE_XML)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const RANGE = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }

async function planWith(options: {
  filterStates?: SheetFilterState[]
  tableEdits?: SheetTableEditRequest[]
  slicerAdditions?: SheetSlicerAddition[]
}) {
  const source = await createBufferEntrySource(await salesWorkbook())
  return planCellEditsToXlsx(
    source,
    [], // edits
    [], // structuralOps
    [], // chartEdits
    undefined, // sheetPlan
    options.filterStates ?? [], // filterStates
    [], // hyperlinkEdits
    [], // cfStates
    [], // dvStates
    [], // sheetProtections
    null, // definedNamesState
    [], // visualAdditions
    [], // pageSetupStates
    [], // noteStates
    [], // tableAdditions
    [], // pivotAdditions
    [], // pivotCacheRefreshPaths
    [], // pivotRefreshUpdates
    [], // visualEdits
    [], // sparklineAdditions
    [], // formulaValues
    null, // themeState
    null, // workbookProtectionState
    [], // protectedRangeStates
    [], // bulkConstantFills
    [], // threadedCommentStates
    options.tableEdits ?? [], // tableEdits
    options.slicerAdditions ?? [], // slicerAdditions
  )
}

describe('slicer × rename (BUG-1751)', () => {
  it('renames the table and still writes the slicer parts and the table filter', async () => {
    const plan = await planWith({
      // The renderer resolves names before the request: both channels carry
      // the post-rename name (recorded under the creation-time name).
      tableEdits: [{ sheetName: 'Data', tableName: 'Sales', rename: 'Sales2026' }],
      slicerAdditions: [{ sheetName: 'Data', tableName: 'Sales2026', colId: 0 }],
      filterStates: [
        {
          sheetName: 'Data',
          tableName: 'Sales2026',
          filter: { range: RANGE, columns: [{ colId: 0, values: ['beta'] }] },
          hiddenRows: [1, 3],
          visibilityRange: RANGE,
        },
      ],
    })

    const tableXml = plan.replaced.get('xl/tables/table1.xml')
    expect(tableXml).toContain('displayName="Sales2026"')
    expect(tableXml).toContain('name="Sales2026"')
    // The table part keeps its autoFilter ref and gains the criteria column.
    expect(tableXml).toContain('<autoFilter ref="A1:B4">')
    expect(tableXml).toContain('<filterColumn colId="0"><filters><filter val="beta"/>')

    expect(plan.added.get('xl/slicers/slicer1.xml')).toBeDefined()
    const cacheXml = plan.added.get('xl/slicerCaches/slicerCache1.xml')
    expect(cacheXml).toContain('<x15:tableSlicerCache tableId="1" column="1"/>')

    const worksheetXml = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(worksheetXml).toContain('<row r="2" hidden="1">')
    expect(worksheetXml).toContain('<row r="4" hidden="1">')
    // No worksheet-level autoFilter over the table: criteria live in the part.
    expect(worksheetXml.startsWith('<?xml')).toBe(true)
    expect(worksheetXml).not.toContain('<autoFilter')
  })

  it('fails closed when an unresolved entry still carries the pre-rename name', async () => {
    // The renderer prevents this shape; the gateway's ordering (tableEdits
    // before slicer additions) is what makes the rename observable at all.
    await expect(
      planWith({
        tableEdits: [{ sheetName: 'Data', tableName: 'Sales', rename: 'Sales2026' }],
        slicerAdditions: [{ sheetName: 'Data', tableName: 'Sales', colId: 0 }],
      }),
    ).rejects.toThrow('Table "Sales" was not found on this sheet.')
  })
})

describe('slicer × convert to range (BUG-1751)', () => {
  it('removes the table with its autoFilter and keeps the filter rows hidden', async () => {
    const plan = await planWith({
      // The criteria snapshot Excel semantics produce: no table part to own
      // criteria anymore, the filtered rows stay hidden as plain row state.
      tableEdits: [{ sheetName: 'Data', tableName: 'Sales', convertToRange: true }],
      slicerAdditions: [],
      filterStates: [
        {
          sheetName: 'Data',
          filter: null,
          hiddenRows: [1, 3],
          visibilityRange: RANGE,
        },
      ],
    })

    expect(plan.removedEntries).toContain('xl/tables/table1.xml')
    const worksheetXml = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(worksheetXml).not.toContain('<tableParts')
    expect(worksheetXml).not.toContain('<autoFilter')
    expect(worksheetXml).toContain('<row r="2" hidden="1">')
    expect(worksheetXml).toContain('<row r="4" hidden="1">')
    // The formulas the rewrite owns are untouched here (none in the fixture);
    // the data rows survive verbatim.
    expect(worksheetXml).toContain('<t>beta</t>')
  })
})

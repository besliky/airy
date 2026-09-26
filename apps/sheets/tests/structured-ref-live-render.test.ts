import { existsSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterAll, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { z } from 'zod'

import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { patchWorksheetRangeInner, runStructuredRefRecalc } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { ICellData } from '@univerjs/core'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

/// BUG-1749 (audit PS-1): structured references were dead in a live session.
/// The import channel collapsed formula cells to style-only (`{"s":"…"}`),
/// and typing `=SUM(Sales[Amount])` stuck at #NAME? — while the sidecar
/// recalc channel computed the same workbook correctly (11.5/100/200). The
/// renderer must keep the formula in the grid (install channel) and fill the
/// values from the sidecar engine (overlay pass), for file cells and typed
/// formulas alike. This file drives the RENDERER channel against the real
/// sidecar binary and a real table fixture — the live path, not just the
/// engine-level probe.

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

const recalcResultSchema = z.object({
  cells: z.array(
    z.object({
      sheet: z.string(),
      row: z.number(),
      column: z.number(),
      formatted: z.string(),
      number: z.number().optional(),
      isFormula: z.boolean(),
    }),
  ),
  cached: z.boolean(),
})

const SREF_FORMULAS: Record<string, string> = {
  // E2 (row 1, column 4): this-row shorthand across two tables → 10 + 1.5.
  '1:4': '=Sales[@Amount]+Prices[@[Unit Price]]',
  // E7 (row 6, column 4): the data body only → 10+20+30+40.
  '6:4': '=SUM(Sales[Amount])',
  // E8 (row 7, column 4): #All spans headers + data + totals → 100 + 100.
  '7:4': '=SUM(Sales[#All])',
}

// Same Sales/Prices tables as the engine-channel fixture
// (xlsx-table-structured-refs.test.ts), plus a formula column E with NO
// cached values — the audit corpus shape Excel recomputes on open.
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
      '<dimension ref="A1:E8"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Product</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Unit Price</t></is></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>alpha</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>1.5</v></c><c r="E2"><f>Sales[@Amount]+Prices[@[Unit Price]]</f></c></row>' +
      '<row r="3"><c r="A3" t="inlineStr"><is><t>beta</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>2.5</v></c></row>' +
      '<row r="4"><c r="A4" t="inlineStr"><is><t>gamma</t></is></c><c r="B4"><v>30</v></c><c r="C4"><v>3.5</v></c></row>' +
      '<row r="5"><c r="A5" t="inlineStr"><is><t>delta</t></is></c><c r="B5"><v>40</v></c><c r="C5"><v>4.5</v></c></row>' +
      '<row r="6"><c r="A6" t="inlineStr"><is><t>Total</t></is></c><c r="B6"><v>100</v></c></row>' +
      '<row r="7"><c r="E7"><f>SUM(Sales[Amount])</f></c></row>' +
      '<row r="8"><c r="E8"><f>SUM(Sales[#All])</f></c></row>' +
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
      '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>' +
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
  const directory = await mkdtemp(join(tmpdir(), 'sref-live-render-'))
  cleanups.push(directory)
  const path = join(directory, name)
  await writeFile(path, await buildTableFixture())
  return path
}

describe('structured-reference import install', () => {
  // Regression for the audit's import symptom: E-column cells arrived in the
  // grid as style-only `{"s":"…"}` because containsUnresolvedNames counted
  // the table token as an unresolved name → formulaKeepsCache "always" →
  // the formula was dropped entirely.
  it('keeps structured-reference formulas in the grid instead of collapsing them', () => {
    let captured: ICellData[][] = []
    const worksheet = {
      getRange: () => ({
        setValues: (matrix: ICellData[][]) => {
          captured = matrix
        },
      }),
    }
    patchWorksheetRangeInner(
      worksheet as never,
      undefined,
      { startRow: 0, endRow: 7, startColumn: 0, endColumn: 4 },
      [
        { row: 1, column: 4, value: null, formula: SREF_FORMULAS['1:4'] },
        { row: 6, column: 4, value: null, formula: SREF_FORMULAS['6:4'] },
        // File cache present (Excel-authored): the cached value must survive
        // next to the formula instead of a bare value-only install.
        { row: 7, column: 4, value: 200, formula: SREF_FORMULAS['7:4'] },
        // Control: a defined name stays cache-kept (no compute channel).
        { row: 6, column: 0, value: null, formula: '=Total*2' },
      ],
      [],
      [],
      [],
      null,
      true,
    )
    expect(captured[1]?.[4]).toEqual({ f: '=Sales[@Amount]+Prices[@[Unit Price]]' })
    expect(captured[6]?.[4]).toEqual({ f: '=SUM(Sales[Amount])' })
    expect(captured[7]?.[4]).toEqual({ f: '=SUM(Sales[#All])', v: 200 })
    expect(captured[6]?.[0]).toEqual({ v: null })
  })
})

type InstalledCell = { row: number; column: number; matrix: ICellData[][] }

function makeStructuredRefState(overrides?: Partial<Record<string, unknown>>): LazyWorkbookState {
  const journal = createEditJournal()
  const state = {
    formulaMode: true,
    file: { sessionId: 'session-1' },
    editJournal: journal,
    formulaText: new Map([['sheet-1', new Map(Object.entries(SREF_FORMULAS))]]),
    loadedRanges: new Map([['sheet-1', { startRow: 0, endRow: 20, startColumn: 0, endColumn: 6 }]]),
    recalc: {
      timer: null,
      generation: 0,
      failures: 0,
      engineOverBudget: false,
      formulaCells: new Map(),
      structuredRefCells: new Map([['sheet-1', new Set(Object.keys(SREF_FORMULAS))]]),
      overlay: new Map(),
      follow: new Map(),
      running: false,
      lastRunAt: 0,
    },
    ...overrides,
  }
  return state as unknown as LazyWorkbookState
}

describe('structured references through the live renderer channel', { skip: !HAS_SIDECAR }, () => {
  it('fills import cells and typed formulas from the sidecar (import → read → input → recalc)', async () => {
    const path = await writeFixture('sales.xlsx')
    const client = new XlsxSidecarClient(SIDECAR)
    const previousWindow = (globalThis as { window?: unknown }).window
    try {
      // Mirror the app wire (sheets-main recalcWorkbook): sheet ids to file
      // names on the way in, engine cells tagged back with the sheet id.
      ;(globalThis as { window?: unknown }).window = {
        desktopApi: {
          recalcWorkbook: async (request: {
            edits: { sheetId: string; row: number; column: number; input: string }[]
            reads: {
              sheetId: string
              range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
            }[]
          }) => {
            const result = recalcResultSchema.parse(
              await client.recalcCells({
                path,
                edits: request.edits.map((edit) => ({
                  sheet: 'Data',
                  row: edit.row,
                  column: edit.column,
                  input: edit.input,
                })),
                reads: request.reads.map((read) => ({
                  sheet: 'Data',
                  range: read.range,
                })),
              }),
            )
            return {
              cells: result.cells.map((cell) => ({ ...cell, sheetId: 'sheet-1' })),
              cached: result.cached,
            }
          },
        },
      }

      // The user types =SUM(Sales[Amount]) into E11 (row 10) — the audit's
      // repro cell. In the grid this evaluates to #NAME? (no table registry);
      // the journal input is what the sidecar must compute from.
      const state = makeStructuredRefState()
      recordSetRangeValues(state.editJournal, 'sheet-1', {
        10: { 4: { f: '=SUM(Sales[Amount])' } },
      })

      const installed: InstalledCell[] = []
      const worksheet = {
        getSheetId: () => 'sheet-1',
        getRange: (row: number, column: number, _rows: number, _columns: number) => ({
          setValues: (matrix: ICellData[][]) => {
            installed.push({ row, column, matrix })
          },
        }),
      }
      const runtime = {
        univerAPI: {
          getActiveWorkbook: () => ({
            getActiveSheet: () => ({ getSheetId: () => 'sheet-1' }),
            getSheetBySheetId: () => worksheet,
          }),
        },
      }

      await runStructuredRefRecalc(runtime as never, { current: state }, state, () => undefined)

      const overlay = state.recalc.overlay.get('sheet-1')
      // Import channel: file formulas computed by the engine.
      expect(overlay?.get('1:4')?.v).toBe(11.5)
      expect(overlay?.get('6:4')?.v).toBe(100)
      expect(overlay?.get('7:4')?.v).toBe(200)
      // Input channel: the typed structured reference resolved, not #NAME?.
      expect(overlay?.get('10:4')?.v).toBe(100)
      // All four values were pinned onto the grid.
      expect(installed.length).toBeGreaterThanOrEqual(4)
    } finally {
      ;(globalThis as { window?: unknown }).window = previousWindow
      client.stop()
    }
  })

  it('round-trips the fixture through openpyxl with the table intact', async () => {
    if (!HAS_OPENPYXL) return
    const path = await writeFixture('sales-openpyxl.xlsx')
    const sheetXml = execFileSync(PYTHON as string, [
      '-c',
      [
        'import openpyxl, sys',
        'wb = openpyxl.load_workbook(sys.argv[1])',
        'ws = wb["Data"]',
        'print(ws.cell(row=2, column=5).value)',
        'print(ws.cell(row=7, column=5).value)',
      ].join('\n'),
      path,
    ])
      .toString()
      .trim()
      .split('\n')
    // openpyxl reads the formula text: the file carries live structured refs.
    expect(sheetXml[0]).toBe('=Sales[@Amount]+Prices[@[Unit Price]]')
    expect(sheetXml[1]).toBe('=SUM(Sales[Amount])')
  })
})

// In-memory stub of the sidecar XlsxIo surface: enough protocol behavior for
// the xlsx session unit tests (open metadata, canned read_range results,
// convert_workbook that writes a placeholder .xlsx) with no binary required.
// The save module is mocked separately per test file (vi.mock) because its
// real implementation drives the whole gateway planning pipeline.
import { writeFile } from 'node:fs/promises'

import type { XlsxIo } from '../../src/xlsx/sidecar-client.js'

export interface StubSheet {
  readonly id: string
  readonly name: string
  readonly rowCount: number
  readonly columnCount: number
}

export interface StubRangeCell {
  readonly row: number
  readonly column: number
  readonly value?: unknown
  readonly formula?: string
}

/** canned read_formula_cells reply cell (the file's indexed formula cells) */
export interface StubFormulaCell {
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly value?: unknown
}

/** canned recalc_cells reply cell (the engine's evaluated values) */
export interface StubRecalcCell {
  readonly sheet: string
  readonly row: number
  readonly column: number
  readonly number?: number
  readonly isFormula: boolean
}

export interface StubIoOptions {
  sheets?: readonly StubSheet[]
  cells?: readonly StubRangeCell[]
  /** per-sheet formula-cell listing for read_formula_cells */
  formulaCells?: readonly StubFormulaCell[]
  /** read_formula_cells reply flags (defaults: complete and not truncated) */
  formulaIndexingComplete?: boolean
  /**
   * read_formula_cells reports an incomplete index for this many polls per
   * sheet before completing — models the sidecar's lazy background indexer,
   * which the first call spawns and later calls observe (BUG-1776). Undefined
   * keeps every reply at formulaIndexingComplete.
   */
  formulaIndexingCompleteAfter?: number
  formulaTruncated?: boolean
  /** evaluated cells recalc_cells reports (and optionally the error it raises) */
  recalcCells?: readonly StubRecalcCell[]
  recalcError?: Error
  /** error the stub restamp_recalc raises (simulates a pre-PERF-1778 sidecar) */
  restampError?: Error
  openError?: Error
  /** raw byte length the stub open reply reports (real sidecars always
   * send it; omitted → the reply has no rawBytes, the pre-BUG-1305 shape) */
  openRawBytes?: number
  /** raw source byte length the stub convert reply reports (same wire
   * contract; omitted → no sourceBytes in the reply) */
  convertSourceBytes?: number
}

export interface StubIo extends XlsxIo {
  /** calls per session id: open -> readRange/close counts for assertions */
  readonly calls: {
    open: string[]
    readRange: string[]
    close: string[]
    convert: string[]
    readFormulaCells: string[]
    recalcCells: string[]
    restampRecalc: string[]
  }
  /** every read_range request's range as "r0..r1 x c0..c1" (0-based, inclusive) */
  readonly readRanges: string[]
  /** every recalc_cells request (path + edits + read count) for assertions */
  readonly recalcRequests: Array<{
    path: string
    edits: readonly { sheet: string; row: number; column: number; input: string }[]
    reads: readonly { sheet: string; range: { startRow: number; endRow: number } }[]
  }>
}

export function makeStubIo(options: StubIoOptions = {}): StubIo {
  const sheets = options.sheets ?? [
    { id: 'sheet-0', name: 'Sheet1', rowCount: 100, columnCount: 26 },
    { id: 'sheet-1', name: 'Data', rowCount: 50, columnCount: 10 },
  ]
  const calls = {
    open: [] as string[],
    readRange: [] as string[],
    close: [] as string[],
    convert: [] as string[],
    readFormulaCells: [] as string[],
    recalcCells: [] as string[],
    restampRecalc: [] as string[],
  }
  const readRanges: string[] = []
  const recalcRequests: StubIo['recalcRequests'] = []
  const formulaCellPolls = new Map<string, number>()
  let sessionCounter = 0
  const io: StubIo = {
    calls,
    readRanges,
    recalcRequests,
    async open(path: string) {
      calls.open.push(path)
      if (options.openError) throw options.openError
      sessionCounter += 1
      return {
        sessionId: `stub-session-${String(sessionCounter)}`,
        name: 'stub',
        entryCount: 2,
        sheets: sheets.map((sheet, index) => ({ ...sheet, index })),
        activeTab: 0,
        ...(options.openRawBytes !== undefined ? { rawBytes: options.openRawBytes } : {}),
      }
    },
    async readRange(input: {
      sessionId: string
      sheetId: string
      range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
    }) {
      calls.readRange.push(`${input.sessionId}:${input.sheetId}`)
      readRanges.push(
        `${String(input.range.startRow)}..${String(input.range.endRow)}` +
          ` x ${String(input.range.startColumn)}..${String(input.range.endColumn)}`,
      )
      const knownSheet = sheets.some((candidate) => candidate.id === input.sheetId)
      return {
        cells: knownSheet ? [...(options.cells ?? [])] : [],
        indexingComplete: true,
      }
    },
    async readFormulaCells(input: { sessionId: string; sheetId: string }) {
      calls.readFormulaCells.push(`${input.sessionId}:${input.sheetId}`)
      const polls = formulaCellPolls.get(input.sheetId) ?? 0
      formulaCellPolls.set(input.sheetId, polls + 1)
      const lazyAfter = options.formulaIndexingCompleteAfter
      const complete =
        lazyAfter === undefined ? options.formulaIndexingComplete !== false : polls >= lazyAfter
      return {
        cells: (options.formulaCells ?? []).filter((cell) => cell.sheetId === input.sheetId),
        indexingComplete: complete,
        truncated: options.formulaTruncated === true,
      }
    },
    async recalcCells(input: {
      path: string
      edits: readonly { sheet: string; row: number; column: number; input: string }[]
      reads: readonly {
        sheet: string
        range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
      }[]
    }) {
      calls.recalcCells.push(input.path)
      recalcRequests.push({
        path: input.path,
        edits: input.edits.map((edit) => ({ ...edit })),
        reads: input.reads.map((read) => ({
          sheet: read.sheet,
          range: { startRow: read.range.startRow, endRow: read.range.endRow },
        })),
      })
      if (options.recalcError) throw options.recalcError
      return { cells: [...(options.recalcCells ?? [])] }
    },
    async restampRecalc(path: string) {
      calls.restampRecalc.push(path)
      if (options.restampError) throw options.restampError
      return { restamped: true }
    },
    async close(sessionId: string) {
      calls.close.push(sessionId)
    },
    async convertWorkbook(input: { path: string; targetPath: string }) {
      calls.convert.push(`${input.path}->${input.targetPath}`)
      // a real conversion writes a workbook; the stub writes placeholder
      // bytes (nothing parses them while the save module is mocked)
      await writeFile(input.targetPath, 'stub-xlsx-bytes')
      return {
        sheets: 1,
        cells: 0,
        ...(options.convertSourceBytes !== undefined
          ? { sourceBytes: options.convertSourceBytes }
          : {}),
      }
    },
    async archiveManifest() {
      return { entries: [] }
    },
    async readEntries() {
      return { entries: [] }
    },
    async scanEntries() {
      return { matches: [] }
    },
    async saveArchive() {
      return { beforeEntries: [], afterEntries: [] }
    },
  }
  return io
}

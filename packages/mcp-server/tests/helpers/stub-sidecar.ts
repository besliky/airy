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

export interface StubIoOptions {
  sheets?: readonly StubSheet[]
  cells?: readonly StubRangeCell[]
  openError?: Error
}

export interface StubIo extends XlsxIo {
  /** calls per session id: open -> readRange/close counts for assertions */
  readonly calls: { open: string[]; readRange: string[]; close: string[]; convert: string[] }
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
  }
  let sessionCounter = 0
  const io: StubIo = {
    calls,
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
      }
    },
    async readRange(input: { sessionId: string; sheetId: string }) {
      calls.readRange.push(`${input.sessionId}:${input.sheetId}`)
      const knownSheet = sheets.some((candidate) => candidate.id === input.sheetId)
      return {
        cells: knownSheet ? [...(options.cells ?? [])] : [],
        indexingComplete: true,
      }
    },
    async readFormulaCells() {
      return { cells: [], indexingComplete: true, truncated: false }
    },
    async close(sessionId: string) {
      calls.close.push(sessionId)
    },
    async convertWorkbook(input: { path: string; targetPath: string }) {
      calls.convert.push(`${input.path}->${input.targetPath}`)
      // a real conversion writes a workbook; the stub writes placeholder
      // bytes (nothing parses them while the save module is mocked)
      await writeFile(input.targetPath, 'stub-xlsx-bytes')
      return { sheets: 1, cells: 0 }
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

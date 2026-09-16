// Headless xlsx session (Phase 3): the sidecar owns reading (lazy-indexed
// ranges, values and formulas) and archive reassembly; this session keeps the
// agent-facing state — sheet index, a pending cell-edit journal, the legacy
// origin when the book was imported via conversion — and drives the app's
// streaming save gateway for byte-preserving writes.
//
//   open:   .xlsx/.xlsm -> sidecar open; .xls/.ods -> convert_workbook into a
//           temp .xlsx (calamine path; styles are NOT carried over) then open
//   read:   sidecar read_range (A1-notation ranges, values + formulas)
//   save:   saveWorkbookViaSidecar (apps/sheets gateway) with the edit
//           journal; untouched zip entries are raw-copied byte-identical
//   close:  sidecar close + temp cleanup
import { randomUUID } from 'node:crypto'
import { copyFile, mkdtemp, rename, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, join } from 'node:path'

import { FencingError, assertSaveTargetFree } from '../docx/session.js'
import { resolveConfined, workspaceRoot } from '../docx/paths.js'
import {
  convertViaSoffice,
  findSoffice,
  SOFFICE_FILTERS,
  sofficeMissingError,
} from '../import/soffice.js'
import {
  saveWorkbookViaSidecar,
  type CellEdit,
  type WorkbookRichRun,
  type WorkbookStyleEdit,
} from './save.js'
import type { XlsxIo } from './sidecar-client.js'

// ---- limits (kept inside the ~30k character MCP answer budget) ----

const READ_MAX_CHARS = 30_000
const MAX_READ_CELLS = 20_000

export type WorkbookFormat = 'xlsx' | 'xlsm' | 'xls' | 'ods'

export interface XlsxSheetSummary {
  readonly index: number
  readonly name: string
  readonly id: string
  readonly rowCount: number
  readonly columnCount: number
}

export interface XlsxSessionMeta {
  readonly handle: string
  readonly kind: 'xlsx'
  /** absolute path of the file the agent opened (origin for imports) */
  readonly path: string
  readonly format: WorkbookFormat
  readonly converted: boolean
  readonly editable: boolean
  readonly warnings: readonly string[]
  readonly sheets: readonly XlsxSheetSummary[]
  readonly activeSheetIndex: number
  readonly dirty: boolean
}

export interface XlsxSaveResult {
  readonly path: string
  readonly bytes: number
  readonly format: 'xlsx' | 'ods'
  readonly unchanged: boolean
  readonly warnings: readonly string[]
  readonly touchedEntries: readonly string[]
}

export interface ReadWorkbookOptions {
  /** sheet name or 0-based index; default: overview of all sheets */
  readonly sheet?: string | number
  /** A1 notation range within the sheet ("A1:C10", "B2"); requires sheet */
  readonly range?: string
}

interface FileStamp {
  readonly mtimeMs: number
  readonly size: number
}

interface SidecarOpenInfo {
  readonly sessionId: string
  readonly sheets: readonly XlsxSheetSummary[]
  readonly activeTab: number
}

const CONVERSION_WARNING: Record<'xls' | 'ods', string> = {
  xls:
    'Imported from legacy .xls via conversion: values, formulas and date formats survive, ' +
    'but styling (fonts, fills, borders, widths, merges, charts) is lost on import.',
  ods:
    'Imported from .ods via conversion: values and formulas survive, but ODF styling and ' +
    'features are lost on import.',
}

// ---- small wire parsers (the sidecar returns JSON; validate the fields we rely on) ----

function asRecord(value: unknown, what: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`XLSX sidecar returned a malformed ${what}.`)
  }
  return value as Record<string, unknown>
}

function parseOpenResult(raw: unknown): SidecarOpenInfo {
  const result = asRecord(raw, 'open result')
  const sessionId = result.sessionId
  if (typeof sessionId !== 'string' || sessionId === '') {
    throw new Error('XLSX sidecar returned no sessionId.')
  }
  const sheetsRaw = result.sheets
  if (!Array.isArray(sheetsRaw)) throw new Error('XLSX sidecar returned no sheet list.')
  const sheets: XlsxSheetSummary[] = sheetsRaw.map((sheetRaw, index) => {
    const sheet = asRecord(sheetRaw, 'sheet entry')
    if (typeof sheet.name !== 'string' || typeof sheet.id !== 'string') {
      throw new Error('XLSX sidecar returned a malformed sheet entry.')
    }
    return {
      index,
      name: sheet.name,
      id: sheet.id,
      rowCount: numberOr(sheet.rowCount, 0),
      columnCount: numberOr(sheet.columnCount, 0),
    }
  })
  return { sessionId, sheets, activeTab: numberOr(result.activeTab, 0) }
}

function numberOr(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback
}

interface RangeCell {
  readonly row: number
  readonly column: number
  readonly value?: unknown
  readonly formula?: string
}

function parseRangeResult(raw: unknown): {
  cells: readonly RangeCell[]
  indexingComplete: boolean
} {
  const result = asRecord(raw, 'read_range result')
  const cellsRaw = result.cells
  if (!Array.isArray(cellsRaw)) throw new Error('XLSX sidecar returned no cell list.')
  const cells = cellsRaw.map((cellRaw) => {
    const cell = asRecord(cellRaw, 'cell entry')
    return {
      row: numberOr(cell.row, -1),
      column: numberOr(cell.column, -1),
      ...(cell.value === undefined ? {} : { value: cell.value }),
      ...(typeof cell.formula === 'string' ? { formula: cell.formula } : {}),
    }
  })
  return { cells, indexingComplete: result.indexingComplete === true }
}

// ---- A1 notation ----

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/
const RANGE_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7}):\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/

export function columnFromLabel(label: string): number {
  let value = 0
  for (const char of label.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64)
  return value - 1
}

export function columnToLabel(index: number): string {
  let label = ''
  let value = index + 1
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

/** Parse "B2" or "A1:C10" into an inclusive 0-based range (unordered refs allowed). */
export function parseA1Range(spec: string): {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
} {
  const single = CELL_RE.exec(spec.trim())
  if (single) {
    const row = Number(single[2]) - 1
    const column = columnFromLabel(single[1]!)
    return { startRow: row, endRow: row, startColumn: column, endColumn: column }
  }
  const range = RANGE_RE.exec(spec.trim())
  if (!range) throw new Error(`Invalid A1-style range "${spec}" (expected e.g. "A1:C10" or "B2").`)
  const startColumn = columnFromLabel(range[1]!)
  const endColumn = columnFromLabel(range[3]!)
  const startRow = Number(range[2]) - 1
  const endRow = Number(range[4]) - 1
  return {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startColumn: Math.min(startColumn, endColumn),
    endColumn: Math.max(startColumn, endColumn),
  }
}

// ---- session ----

export class XlsxSession {
  readonly handle: string
  /** absolute path of the file backing the session (temp .xlsx for imports) */
  readonly backingPath: string
  /** user-facing format/path the agent opened */
  readonly format: WorkbookFormat
  readonly originPath: string | null

  private readonly io: XlsxIo
  private readonly root: string
  private readonly tempDir: string | null
  private readonly originStamp: FileStamp | null
  private readonly warnings: string[] = []
  private readonly edits: CellEdit[] = []
  private sessionId: string
  private sheets: readonly XlsxSheetSummary[]
  private activeSheetIndex: number
  private baseline: FileStamp | null
  private savedPath: string | null = null
  /** every target this session has written (repeat save-as needs no overwrite) */
  private readonly savedTargets = new Set<string>()

  private constructor(options: {
    handle: string
    backingPath: string
    format: WorkbookFormat
    originPath: string | null
    io: XlsxIo
    root: string
    tempDir: string | null
    originStamp: FileStamp | null
    warnings: string[]
    open: SidecarOpenInfo
    baseline: FileStamp | null
  }) {
    this.handle = options.handle
    this.backingPath = options.backingPath
    this.format = options.format
    this.originPath = options.originPath
    this.io = options.io
    this.root = options.root
    this.tempDir = options.tempDir
    this.originStamp = options.originStamp
    this.warnings = [...options.warnings]
    this.sessionId = options.open.sessionId
    this.sheets = options.open.sheets
    this.activeSheetIndex = options.open.activeTab
    this.baseline = options.baseline
  }

  /** Open a workbook inside the workspace root; .xls/.ods convert to temp .xlsx first. */
  static async open(rawPath: string, root?: string, io?: XlsxIo): Promise<XlsxSession> {
    const path = resolveConfined(rawPath, root)
    const format = extensionFormat(path)
    if (format === null) {
      throw new Error(
        `Unsupported workbook extension "${extname(path) || '(none)'}" for "${path}" ` +
          '(expected .xlsx, .xlsm, .xls or .ods).',
      )
    }
    const client = io ?? (await defaultIo())
    let backingPath = path
    let tempDir: string | null = null
    const warnings: string[] = []
    if (format === 'xls' || format === 'ods') {
      tempDir = await mkdtemp(join(tmpdir(), 'airy-import-'))
      backingPath = join(tempDir, `${siblingStem(basename(path))}.xlsx`)
      try {
        await client.convertWorkbook({ path, targetPath: backingPath })
      } catch (e) {
        await rm(tempDir, { recursive: true, force: true })
        throw new Error(
          `Cannot import "${path}" as .xlsx: ${e instanceof Error ? e.message : String(e)}`,
          { cause: e },
        )
      }
      warnings.push(CONVERSION_WARNING[format])
    }

    try {
      const openInfo = parseOpenResult(await client.open(backingPath))
      const baseline = await statOrNull(backingPath)
      return new XlsxSession({
        handle: randomUUID(),
        backingPath,
        format,
        originPath: format === 'xls' || format === 'ods' ? path : null,
        io: client,
        root: root ?? workspaceRoot(),
        tempDir,
        originStamp: format === 'xls' || format === 'ods' ? await statOrNull(path) : null,
        warnings,
        open: openInfo,
        baseline,
      })
    } catch (e) {
      if (tempDir) await rm(tempDir, { recursive: true, force: true })
      throw e
    }
  }

  // ---- reading ----

  meta(): XlsxSessionMeta {
    return {
      handle: this.handle,
      kind: 'xlsx',
      path: this.originPath ?? this.backingPath,
      format: this.format,
      converted: this.originPath !== null,
      editable: true,
      warnings: [...this.warnings],
      sheets: this.sheets.map((sheet) => ({ ...sheet })),
      activeSheetIndex: this.activeSheetIndex,
      dirty: this.isDirty(),
    }
  }

  private isDirty(): boolean {
    return this.edits.length > 0
  }

  private resolveSheet(sheet: string | number): XlsxSheetSummary {
    const found =
      typeof sheet === 'number'
        ? this.sheets[sheet]
        : this.sheets.find((candidate) => candidate.name === sheet)
    if (!found) {
      throw new Error(
        `No sheet ${typeof sheet === 'number' ? `index ${String(sheet)}` : `"${sheet}"`} ` +
          `(workbook has ${String(this.sheets.length)}: ${this.sheetNames().join(', ')}).`,
      )
    }
    return found
  }

  private sheetNames(): string[] {
    return this.sheets.map((sheet) => sheet.name)
  }

  /**
   * Agent-facing read: without options the sheet overview
   * (`index|name|id|rows x cols`), with sheet (+range) a pipe table of cell
   * values with formulas inline. Output stays under ~30k characters.
   */
  async readWorkbook(options: ReadWorkbookOptions = {}): Promise<string> {
    if (options.sheet === undefined) {
      if (options.range !== undefined) {
        throw new Error('A range requires the sheet it belongs to.')
      }
      const lines = this.sheets.map(
        (sheet) =>
          `${String(sheet.index)}|${sheet.name}|${sheet.id}|${sheet.rowCount} x ${String(sheet.columnCount)}`,
      )
      return [
        `The workbook has ${String(this.sheets.length)} sheet(s) (index|name|id|rows x cols):`,
        ...lines,
        'Read cells with read_workbook: pass sheet (name or index) and an A1-style range, e.g. sheet "Sheet1", range "A1:E10".',
        ...(this.warnings.length > 0
          ? ['', 'Warnings:', ...this.warnings.map((w) => `- ${w}`)]
          : []),
      ].join('\n')
    }

    const sheet = this.resolveSheet(options.sheet)
    if (options.range === undefined) {
      // whole used area is often huge; start the agent with the top corner
      const rows = Math.min(sheet.rowCount, 20)
      const columns = Math.min(sheet.columnCount, 10)
      if (rows <= 0 || columns <= 0) {
        return `Sheet "${sheet.name}" is empty.`
      }
      return this.readRange(sheet, {
        startRow: 0,
        endRow: rows - 1,
        startColumn: 0,
        endColumn: columns - 1,
      })
    }
    return this.readRange(sheet, parseA1Range(options.range))
  }

  private async readRange(
    sheet: XlsxSheetSummary,
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  ): Promise<string> {
    const clamped = {
      startRow: Math.max(0, range.startRow),
      endRow: Math.min(range.endRow, sheet.rowCount - 1),
      startColumn: Math.max(0, range.startColumn),
      endColumn: Math.min(range.endColumn, sheet.columnCount - 1),
    }
    if (clamped.endRow < clamped.startRow || clamped.endColumn < clamped.startColumn) {
      throw new Error(
        `Range is outside sheet "${sheet.name}" (${String(sheet.rowCount)} rows x ${String(sheet.columnCount)} columns).`,
      )
    }
    const cells =
      (clamped.endRow - clamped.startRow + 1) * (clamped.endColumn - clamped.startColumn + 1)
    if (cells > MAX_READ_CELLS) {
      throw new Error(
        `Range has ${String(cells)} cells (limit ${String(MAX_READ_CELLS)}). Split it into smaller reads.`,
      )
    }
    const result = parseRangeResult(
      await this.io.readRange({ sessionId: this.sessionId, sheetId: sheet.id, range: clamped }),
    )
    const cellByCoordinate = new Map<string, RangeCell>()
    for (const cell of result.cells)
      cellByCoordinate.set(`${String(cell.row)},${String(cell.column)}`, cell)
    const header = [
      ' ',
      ...Array.from({ length: clamped.endColumn - clamped.startColumn + 1 }, (_, i) =>
        columnToLabel(clamped.startColumn + i),
      ),
    ]
      .map((cell) => `|${cell}`)
      .join('')
      .concat('|')
    const body = Array.from({ length: clamped.endRow - clamped.startRow + 1 }, (_, rowIndex) => {
      const rowNumber = clamped.startRow + rowIndex + 1
      const cellsInRow = Array.from(
        { length: clamped.endColumn - clamped.startColumn + 1 },
        (_, columnIndex) => {
          const cell = cellByCoordinate.get(
            `${String(clamped.startRow + rowIndex)},${String(clamped.startColumn + columnIndex)}`,
          )
          return renderCell(cell)
        },
      )
      return `|${String(rowNumber)}${cellsInRow.map((cell) => `|${cell}`).join('')}|`
    })
    const from = `${columnToLabel(clamped.startColumn)}${String(clamped.startRow + 1)}`
    const to = `${columnToLabel(clamped.endColumn)}${String(clamped.endRow + 1)}`
    const note = result.indexingComplete
      ? ''
      : '\n(Indexing still in progress — the data above may be partial; re-read the range shortly.)'
    const out = [`Sheet "${sheet.name}", range ${from}:${to}:`, header, ...body, note].join('\n')
    return out.length > READ_MAX_CHARS
      ? `${out.slice(0, READ_MAX_CHARS)}\n…(output truncated at ${String(READ_MAX_CHARS)} characters; request a narrower range)`
      : out
  }

  // ---- editing ----

  /**
   * Journal cell edits (patch merge per cell, insertion order kept).
   * Values write as constants; a formula writes `=...` without a cached
   * value, so full-featured spreadsheet apps recalculate it on open (a
   * formula wins over a value when both are given). A style/rich/styleReset
   * patch merges onto the journaled content edit for the same cell — the
   * same present=set patch semantics as the docx ops. dryRun validates the
   * whole batch (sheet names, refs, edit shapes) without journaling.
   */
  setCells(
    input: {
      sheet: string | number
      cells: ReadonlyArray<{
        ref: string
        value?: string | number | boolean | null
        formula?: string
        style?: WorkbookStyleEdit
        rich?: readonly WorkbookRichRun[]
        styleReset?: boolean
      }>
    },
    dryRun = false,
  ): { journaled: number } {
    const sheet = this.resolveSheet(input.sheet)
    let journaled = 0
    for (const cell of input.cells) {
      const address = parseA1Range(cell.ref)
      if (address.startRow !== address.endRow || address.startColumn !== address.endColumn) {
        throw new Error(`Cell ref "${cell.ref}" must be a single cell like "B2", not a range.`)
      }
      const edit = toCellEdit(sheet.name, address.startRow, address.startColumn, cell)
      if (!dryRun) {
        const existing = this.edits.find(
          (candidate) =>
            candidate.sheetName === sheet.name &&
            candidate.row === address.startRow &&
            candidate.column === address.startColumn,
        )
        if (existing) mergeCellEdit(existing, edit)
        else this.edits.push(edit)
      }
      journaled += 1
    }
    return { journaled }
  }

  // ---- saving ----

  /**
   * Save through the streaming gateway: the journal is planned into zip
   * entry patches, the sidecar reassembles the archive (untouched entries
   * raw-copied byte-identical) and the result lands atomically.
   *
   * Default target: the opened .xlsx; for imported .xls/.ods books a fresh
   * sibling .xlsx next to the original (true legacy output is not supported;
   * format 'origin' refuses for .xls and exports .ods via LibreOffice). An
   * explicit target that already exists is refused unless it is the session's
   * own backing file / last output, or overwrite is true.
   */
  async save(
    rawPath?: string,
    format: 'xlsx' | 'origin' = 'xlsx',
    options: { overwrite?: boolean } = {},
  ): Promise<XlsxSaveResult> {
    if (format === 'origin') return this.saveToOrigin()
    const target = resolveConfined(rawPath ?? this.defaultTarget(), this.root)
    await assertSaveTargetFree(
      target,
      [this.backingPath, ...this.savedTargets],
      rawPath,
      options.overwrite,
    )
    if (target === this.backingPath) await this.assertBackingUnchanged()

    const result = await saveWorkbookViaSidecar({
      client: this.io,
      sourcePath: this.backingPath,
      targetPath: target,
      edits: this.edits.map((edit) => ({ ...edit, cell: { ...edit.cell } })),
    })
    const bytes = await statOrNull(target)
    const unchanged = this.edits.length === 0
    this.edits.length = 0
    this.savedPath = target
    this.savedTargets.add(target)
    if (target === this.backingPath) {
      this.baseline = await statOrNull(this.backingPath)
      // the sidecar's in-memory index still reflects the pre-save file:
      // reopen the session so subsequent reads see the saved cells
      await this.reopenSidecar()
    }
    return {
      path: target,
      bytes: bytes?.size ?? 0,
      format: 'xlsx',
      unchanged,
      warnings:
        this.originPath !== null
          ? [
              `Saved as .xlsx; the original ${this.format === 'ods' ? '.ods' : '.xls'} file "${this.originPath}" was left untouched.`,
            ]
          : [],
      touchedEntries: [...result.touchedEntries],
    }
  }

  private defaultTarget(): string {
    if (this.originPath === null) return this.backingPath
    return join(dirname(this.originPath), `${siblingStem(basename(this.originPath))}.xlsx`)
  }

  private async saveToOrigin(): Promise<XlsxSaveResult> {
    if (this.originPath === null || (this.format !== 'xls' && this.format !== 'ods')) {
      throw new Error('format "origin" is only valid for sessions imported from .xls/.ods.')
    }
    if (this.format === 'xls') {
      throw new Error(
        'Writing the legacy .xls format is not supported. Save as .xlsx instead (omit format/path) — ' +
          'the original .xls stays untouched.',
      )
    }
    if (this.originStamp) {
      const current = await statOrNull(this.originPath)
      if (
        !current ||
        current.mtimeMs !== this.originStamp.mtimeMs ||
        current.size !== this.originStamp.size
      ) {
        throw new FencingError(this.originPath)
      }
    }
    // 1. flush the journal into a temp xlsx via the byte-preserving gateway
    const tempXlsx = join(
      await mkdtemp(join(tmpdir(), 'airy-origin-')),
      `${siblingStem(basename(this.originPath))}.xlsx`,
    )
    try {
      await saveWorkbookViaSidecar({
        client: this.io,
        sourcePath: this.backingPath,
        targetPath: tempXlsx,
        edits: this.edits.map((edit) => ({ ...edit, cell: { ...edit.cell } })),
      })
      // 2. export back to .ods through LibreOffice (best-effort fidelity)
      const tool = await findSoffice()
      if (!tool) {
        throw sofficeMissingError('Exporting back to .ods requires LibreOffice.')
      }
      const outDir = dirname(tempXlsx)
      const output = await convertViaSoffice(tool, tempXlsx, {
        filter: SOFFICE_FILTERS.ods,
        extension: 'ods',
        outDir,
      })
      // 3. atomic promote onto the original .ods
      const tmpTarget = join(
        dirname(this.originPath),
        `.${basename(this.originPath)}.airy-${randomUUID()}`,
      )
      await copyFile(output, tmpTarget)
      await rename(tmpTarget, this.originPath)
      const bytes = await statOrNull(this.originPath)
      this.edits.length = 0
      this.savedPath = this.originPath
      this.savedTargets.add(this.originPath)
      return {
        path: this.originPath,
        bytes: bytes?.size ?? 0,
        format: 'ods',
        unchanged: false,
        warnings: [
          'Exported back to .ods via LibreOffice — best-effort fidelity, verify the result.',
        ],
        touchedEntries: [],
      }
    } finally {
      await rm(dirname(tempXlsx), { recursive: true, force: true })
    }
  }

  /**
   * mtime/size fence on the backing file: an in-place save refuses when the
   * file changed on disk since open (the gateway additionally aborts on
   * manifest drift, but this gives the agent the familiar early error).
   */
  private async assertBackingUnchanged(): Promise<void> {
    if (!this.baseline) return
    const current = await statOrNull(this.backingPath)
    if (
      !current ||
      current.mtimeMs !== this.baseline.mtimeMs ||
      current.size !== this.baseline.size
    ) {
      throw new FencingError(this.backingPath)
    }
  }

  private async reopenSidecar(): Promise<void> {
    const previous = this.sessionId
    try {
      await this.io.close(previous)
    } catch {
      // a stale session leak is preferable to failing a completed save
    }
    const openInfo = parseOpenResult(await this.io.open(this.backingPath))
    this.sessionId = openInfo.sessionId
    this.sheets = openInfo.sheets
    this.activeSheetIndex = openInfo.activeTab
  }

  // ---- lifecycle ----

  /** Close the sidecar session and remove the import temp dir when present. */
  async close(): Promise<string[]> {
    const cleaned: string[] = []
    try {
      await this.io.close(this.sessionId)
    } catch {
      // best-effort: the temp dir below still gets swept
    }
    if (this.tempDir) {
      await rm(this.tempDir, { recursive: true, force: true })
      cleaned.push(this.tempDir)
    }
    return cleaned
  }
}

// ---- helpers ----

function extensionFormat(path: string): WorkbookFormat | null {
  const ext = extname(path).replace('.', '').toLowerCase()
  if (ext === 'xlsx' || ext === 'xlsm' || ext === 'xls' || ext === 'ods') return ext
  return null
}

/** "legacy.xls" -> "legacy" (used for converted sibling file names). */
function siblingStem(fileName: string): string {
  const dot = fileName.lastIndexOf('.')
  return dot > 0 ? fileName.slice(0, dot) : fileName
}

async function statOrNull(path: string): Promise<FileStamp | null> {
  try {
    const info = await stat(path)
    return { mtimeMs: info.mtimeMs, size: info.size }
  } catch {
    return null
  }
}

function toCellEdit(
  sheetName: string,
  row: number,
  column: number,
  cell: {
    value?: string | number | boolean | null
    formula?: string
    style?: WorkbookStyleEdit
    rich?: readonly WorkbookRichRun[]
    styleReset?: boolean
  },
): CellEdit {
  if (
    cell.value === undefined &&
    cell.formula === undefined &&
    cell.rich === undefined &&
    cell.style === undefined &&
    cell.styleReset === undefined
  ) {
    throw new Error(
      `The edit for ${sheetName}!${columnToLabel(column)}${String(row + 1)} needs at least one ` +
        'of value, formula, rich, style or styleReset.',
    )
  }
  const stylePatch =
    cell.style === undefined && cell.styleReset === undefined
      ? {}
      : {
          ...(cell.style === undefined ? {} : { style: cell.style }),
          ...(cell.styleReset === undefined ? {} : { styleReset: cell.styleReset }),
        }
  if (cell.formula !== undefined) {
    const formula = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`
    // no cached <v>: apps with a formula engine recalculate on open (an
    // agent-guessed cached result would go stale; value is ignored here)
    return {
      sheetName,
      row,
      column,
      writeValue: true,
      cell: { value: '', formula },
      ...stylePatch,
    }
  }
  if (cell.value !== undefined || cell.rich !== undefined) {
    return {
      sheetName,
      row,
      column,
      writeValue: true,
      cell: { value: cell.value ?? joinedRichText(cell.rich, null) },
      ...(cell.rich === undefined ? {} : { rich: cell.rich }),
      ...stylePatch,
    }
  }
  // style-only: writeValue false keeps the cell's stored content untouched
  return { sheetName, row, column, writeValue: false, cell: { value: null }, ...stylePatch }
}

/** the cell text rich runs spell out (`value` holds the joined text on the wire) */
function joinedRichText(
  rich: readonly WorkbookRichRun[] | undefined,
  fallback: string | number | boolean | null,
): string | number | boolean | null {
  if (rich === undefined || rich.length === 0) return fallback
  return rich.map((run) => run.text).join('')
}

/**
 * Patch-merge a later edit onto a journaled one: a content edit (writeValue)
 * replaces the journaled content (and rich runs), a style patch replaces the
 * journaled style — channels the new edit does not mention stay untouched.
 */
function mergeCellEdit(existing: CellEdit, edit: CellEdit): void {
  const content: Record<string, unknown> = edit.writeValue
    ? { writeValue: true, cell: edit.cell, rich: edit.rich }
    : {}
  Object.assign(
    existing,
    content,
    ...(edit.style === undefined ? [] : [{ style: edit.style }]),
    ...(edit.styleReset === undefined ? [] : [{ styleReset: edit.styleReset }]),
  )
}

function renderCell(cell: RangeCell | undefined): string {
  if (!cell) return ''
  const value =
    cell.value === undefined || cell.value === null
      ? ''
      : typeof cell.value === 'boolean'
        ? cell.value
          ? 'TRUE'
          : 'FALSE'
        : String(cell.value)
  if (cell.formula !== undefined) {
    const formula = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`
    return value === '' ? formula : `${formula} (${value})`
  }
  return value
}

// ---- shared sidecar process ----

let sharedIo: XlsxIo | null = null

/**
 * The process-wide sidecar client: the binary is resolved lazily so merely
 * listing tools (or editing docx) never requires it to exist.
 */
export async function defaultIo(): Promise<XlsxIo> {
  if (sharedIo) return sharedIo
  const { XlsxSidecarClient } = await import('./sidecar-client.js')
  const { findSidecarBinary, sidecarMissingError } = await import('./discovery.js')
  const binary = await findSidecarBinary()
  if (!binary) throw sidecarMissingError()
  sharedIo = new XlsxSidecarClient(binary)
  return sharedIo
}

/** Test seam: run sessions against an injected sidecar client. */
export function setSharedIo(io: XlsxIo | null): void {
  sharedIo = io
}

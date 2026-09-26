// Headless xlsx session (Phase 3): the sidecar owns reading (lazy-indexed
// ranges, values and formulas) and archive reassembly; this session keeps the
// agent-facing state — sheet index, a pending cell-edit journal, the legacy
// origin when the book was imported via conversion — and drives the app's
// streaming save gateway for byte-preserving writes.
//
//   open:   .xlsx/.xlsm -> sidecar open; .xls/.ods -> convert_workbook into a
//           temp .xlsx (calamine path; styles are NOT carried over) then open
//   read:   sidecar read_range (A1-notation ranges, values + formulas) with
//           the pending journal overlaid: read-after-write shows the edits,
//           and journaled cells beyond the on-disk used range are readable
//   save:   saveWorkbookViaSidecar (apps/sheets gateway) with the edit
//           journal; untouched zip entries are raw-copied byte-identical
//   close:  sidecar close + temp cleanup
import { randomUUID } from 'node:crypto'
import { copyFile, mkdtemp, rename, rm, stat } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, join } from 'node:path'

import {
  saveTargetExistsError,
  FencingError,
  asWorkspaceWriteError,
  assertSaveTargetFree,
  withSaveTmpCleanup,
} from '../docx/session.js'
import { assertWorkspaceRootExists, resolveConfined, workspaceRoot } from '../docx/paths.js'
import {
  convertViaSoffice,
  findSoffice,
  SOFFICE_FILTERS,
  sofficeMissingError,
} from '../import/soffice.js'
import { assertWithinOpenCap } from '../sessions/size-fence.js'
import {
  SaveTargetExistsError,
  saveWorkbookViaSidecar,
  type CellEdit,
  type SheetFormulaValues,
  type WorkbookRichRun,
  type WorkbookStyleEdit,
} from './save.js'
// one shared A1 parse/validate for reads and writes (BUG-1632)
import { columnToLabel, parseA1Range } from './refs.js'
import type { XlsxIo } from './sidecar-client.js'

// ---- limits (kept inside the ~30k character MCP answer budget) ----

const READ_MAX_CHARS = 30_000
const MAX_READ_CELLS = 20_000

/**
 * Formula-cache refresh on save (BUG-1761, the MCP twin of the sheets app's
 * recalc overlay from #230): the sidecar's recalc read budget is 20k cells per
 * request — the overlay chunks the formula cells into requests of at most that
 * many 1x1 reads (the journal edits ride every request, so the resident model
 * is reused after the first import).
 */
const MAX_RECALC_READ_CELLS = 20_000
/**
 * Recalculation imports the whole book into the engine; mirrors the sidecar's
 * own speculative-work threshold (PREWARM_MAX_SOURCE_BYTES) so saving a
 * mega-book never triggers a multi-gigabyte model import from a save.
 */
const RECALC_MAX_SOURCE_BYTES = 64_000_000
/** the sidecar refuses recalc requests with more edits (MAX_RECALC_EDITS) */
const MAX_RECALC_EDITS = 10_000
/**
 * The sidecar indexes a worksheet lazily: the first read per sheet spawns a
 * background indexer, and read_formula_cells itself never blocks (the sheets
 * app polls it from its UI), so the cache refresh polls until the index
 * completes. Bounded so a pathological sheet degrades to the honest save
 * warning instead of stalling it (BUG-1776: the overlay used to make exactly
 * one cold call — always indexingComplete:false — and gave up, killing the
 * overlay in the canonical open -> edit -> save flow that never read first).
 */
const RECALC_INDEX_WAIT_MS = 10_000
const RECALC_INDEX_POLL_MS = 250
/** current index wait budget (mutable only through the test seam below) */
let recalcIndexWaitMs = RECALC_INDEX_WAIT_MS

/**
 * Test seam: shrink the index wait budget so unit tests cover the
 * never-completing degradation path without spending real seconds polling
 * (same role as setSharedIo).
 */
export function setRecalcIndexWaitForTests(waitMs: number): void {
  recalcIndexWaitMs = waitMs
}

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
  /** raw byte length the sidecar measured on the file handle it actually
   * opened (BUG-1305); 0 when an older sidecar omits it */
  readonly rawBytes: number
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
  return {
    sessionId,
    sheets,
    activeTab: numberOr(result.activeTab, 0),
    rawBytes: numberOr(result.rawBytes, 0),
  }
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
    // SEC-1102 raw-size fence for parity with the docx/slides opens. The
    // workbook bytes themselves never transit Node — the Rust sidecar reads
    // the file — so this refuses runaway inputs before spawning any work;
    // the declared-uncompressed (zip-bomb) budget for .xlsx runs inside the
    // sidecar itself (SEC-1103: it sums the central-directory sizes and
    // refuses over 1.5 GiB before decompressing), reaching this session as
    // the open reply's error. The same budget also guards the .xls/.ods
    // conversion below (SEC-1301: calamine reads .ods through the same ZIP
    // container, so the sidecar runs the fence before convert_workbook
    // decompresses anything), surfacing inside the "Cannot import" wrap.
    let rawSize = 0
    try {
      rawSize = (await stat(path)).size
    } catch {
      // a missing/unreadable file reports through the open/conversion paths
    }
    assertWithinOpenCap(path, rawSize, 'xlsx')
    const client = io ?? (await defaultIo())
    let backingPath = path
    let tempDir: string | null = null
    const warnings: string[] = []
    if (format === 'xls' || format === 'ods') {
      tempDir = await mkdtemp(join(tmpdir(), 'airy-import-'))
      backingPath = join(tempDir, `${siblingStem(basename(path))}.xlsx`)
      let converted: { sourceBytes?: unknown }
      try {
        converted = (await client.convertWorkbook({ path, targetPath: backingPath })) as {
          sourceBytes?: unknown
        }
      } catch (e) {
        await rm(tempDir, { recursive: true, force: true })
        throw new Error(
          `Cannot import "${path}" as .xlsx: ${e instanceof Error ? e.message : String(e)}`,
          { cause: e },
        )
      }
      // BUG-1305: the stat above ran before the sidecar read the file, so a
      // source that grew past the cap in between slipped the raw fence —
      // re-check against the size the converter itself saw (0 when an older
      // sidecar omits it)
      try {
        assertWithinOpenCap(path, numberOr(converted.sourceBytes, 0), 'xlsx')
      } catch (e) {
        await rm(tempDir, { recursive: true, force: true })
        throw e
      }
      warnings.push(CONVERSION_WARNING[format])
    }

    try {
      const openInfo = parseOpenResult(await client.open(backingPath))
      // BUG-1305: same race for the file the sidecar just opened — re-run
      // the fence against the reply's served size, releasing the sidecar
      // session when the open is refused
      if (openInfo.rawBytes > 0) {
        try {
          assertWithinOpenCap(backingPath, openInfo.rawBytes, 'xlsx')
        } catch (e) {
          await client.close(openInfo.sessionId)
          throw e
        }
      }
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
      const lines = this.sheets.map((sheet) => {
        // journal-aware dims: journaled cells beyond the on-disk used range
        // are readable, so the overview reports the grown area (BUG-1503)
        const extent = this.sheetExtent(sheet)
        return `${String(sheet.index)}|${sheet.name}|${sheet.id}|${String(extent.rowCount)} x ${String(extent.columnCount)}`
      })
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
      // (journal-aware: a fresh cell in an empty sheet is visible there too)
      const extent = this.sheetExtent(sheet)
      const rows = Math.min(extent.rowCount, 20)
      const columns = Math.min(extent.columnCount, 10)
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

  /**
   * Journal-aware sheet bounds: the readable area covers the on-disk used
   * range plus any journaled cell beyond it (BUG-1503 — an edit into a fresh
   * row/column made the cells "not exist" for reads until save).
   */
  private sheetExtent(sheet: XlsxSheetSummary): { rowCount: number; columnCount: number } {
    let rowCount = sheet.rowCount
    let columnCount = sheet.columnCount
    for (const edit of this.edits) {
      if (edit.sheetName !== sheet.name) continue
      rowCount = Math.max(rowCount, edit.row + 1)
      columnCount = Math.max(columnCount, edit.column + 1)
    }
    return { rowCount, columnCount }
  }

  private async readRange(
    sheet: XlsxSheetSummary,
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  ): Promise<string> {
    const extent = this.sheetExtent(sheet)
    const clamped = {
      startRow: Math.max(0, range.startRow),
      endRow: Math.min(range.endRow, extent.rowCount - 1),
      startColumn: Math.max(0, range.startColumn),
      endColumn: Math.min(range.endColumn, extent.columnCount - 1),
    }
    if (clamped.endRow < clamped.startRow || clamped.endColumn < clamped.startColumn) {
      throw new Error(
        `Range is outside sheet "${sheet.name}" (${String(extent.rowCount)} rows x ${String(extent.columnCount)} columns).`,
      )
    }
    const cells =
      (clamped.endRow - clamped.startRow + 1) * (clamped.endColumn - clamped.startColumn + 1)
    if (cells > MAX_READ_CELLS) {
      throw new Error(
        `Range has ${String(cells)} cells (limit ${String(MAX_READ_CELLS)}). Split it into smaller reads.`,
      )
    }
    // the sidecar refuses ranges past the sheet's on-disk used area, so only
    // the intersection is fetched; cells beyond it come from the journal
    const diskEndRow = Math.min(clamped.endRow, sheet.rowCount - 1)
    const diskEndColumn = Math.min(clamped.endColumn, sheet.columnCount - 1)
    let result: { cells: readonly RangeCell[]; indexingComplete: boolean } = {
      cells: [],
      indexingComplete: true,
    }
    if (diskEndRow >= clamped.startRow && diskEndColumn >= clamped.startColumn) {
      result = parseRangeResult(
        await this.io.readRange({
          sessionId: this.sessionId,
          sheetId: sheet.id,
          range: {
            startRow: clamped.startRow,
            endRow: diskEndRow,
            startColumn: clamped.startColumn,
            endColumn: diskEndColumn,
          },
        }),
      )
    }
    const cellByCoordinate = new Map<string, RangeCell>()
    for (const cell of result.cells)
      cellByCoordinate.set(`${String(cell.row)},${String(cell.column)}`, cell)
    // journal overlay: read-after-write must show the pending edits, not the
    // on-disk state (BUG-1503). Content edits replace the cell; style-only
    // edits change nothing the value table shows. A journaled formula has no
    // cached result, so it renders as "=FORMULA" alone.
    let overlaid = 0
    for (const edit of this.edits) {
      if (edit.sheetName !== sheet.name || !edit.writeValue) continue
      if (
        edit.row < clamped.startRow ||
        edit.row > clamped.endRow ||
        edit.column < clamped.startColumn ||
        edit.column > clamped.endColumn
      ) {
        continue
      }
      const key = `${String(edit.row)},${String(edit.column)}`
      if (edit.cell.formula !== undefined) {
        cellByCoordinate.set(key, {
          row: edit.row,
          column: edit.column,
          formula: edit.cell.formula,
        })
      } else if (edit.cell.value === null) {
        // an explicit null clears the cell
        cellByCoordinate.set(key, { row: edit.row, column: edit.column })
      } else {
        cellByCoordinate.set(key, {
          row: edit.row,
          column: edit.column,
          value: edit.cell.value,
        })
      }
      overlaid += 1
    }
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
    const notes: string[] = []
    if (!result.indexingComplete) {
      notes.push(
        '(Indexing still in progress — the data above may be partial; re-read the range shortly.)',
      )
    }
    if (overlaid > 0) {
      notes.push(
        `(${String(overlaid)} pending edit(s) from this session's journal are included; save_document persists them.)`,
      )
    }
    const note = notes.length > 0 ? `\n${notes.join(' ')}` : ''
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
   *
   * The batch is validated in full before anything is journaled (BUG-1632):
   * a bad ref — e.g. "A0", valid A1 syntax whose 0-based row is -1 — used to
   * be journaled and then poisoned every later save ("Invalid cell
   * coordinates: -1,0"). The shared parse/validate in refs.ts refuses such
   * refs up front, so a rejected batch leaves the journal untouched.
   *
   * The return distinguishes the input count (`journaled`: every edit the
   * batch spelled, including later edits to already-journaled cells) from
   * the journal growth (`merged`: entries this call added — several edits to
   * one cell merge into a single journal entry).
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
  ): { journaled: number; merged: number } {
    const sheet = this.resolveSheet(input.sheet)
    // pass 1 — resolve and validate every edit before journaling any, so a
    // per-cell failure (bad sheet, non-addressable/range ref, edit without
    // channels) refuses the whole batch atomically
    const parsed = input.cells.map((cell) => {
      const address = parseA1Range(cell.ref)
      if (address.startRow !== address.endRow || address.startColumn !== address.endColumn) {
        throw new Error(`Cell ref "${cell.ref}" must be a single cell like "B2", not a range.`)
      }
      const row = address.startRow
      const column = address.startColumn
      return { row, column, edit: toCellEdit(sheet.name, row, column, cell) }
    })
    if (dryRun) return { journaled: parsed.length, merged: 0 }
    let merged = 0
    for (const { row, column, edit } of parsed) {
      const existing = this.edits.find(
        (candidate) =>
          candidate.sheetName === sheet.name &&
          candidate.row === row &&
          candidate.column === column,
      )
      if (existing) mergeCellEdit(existing, edit)
      else {
        this.edits.push(edit)
        merged += 1
      }
    }
    return { journaled: parsed.length, merged }
  }

  // ---- saving ----

  /**
   * Save through the streaming gateway: the journal is planned into zip
   * entry patches, the sidecar reassembles the archive (untouched entries
   * raw-copied byte-identical) and the result lands atomically.
   *
   * Stale-root refusal: when the pinned workspace root has been moved/renamed
   * since open, the save fails with StaleWorkspaceRootError (BUG-1103) — the
   * same guard the docx/line/slides sessions run, so a renamed root yields
   * the documented refusal instead of a raw ENOENT from the gateway.
   *
   * Default target: the opened .xlsx; for imported .xls/.ods books a fresh
   * sibling .xlsx next to the original (true legacy output is not supported;
   * format 'origin' refuses for .xls and exports .ods via LibreOffice). A
   * target (explicit or default) that already exists is refused unless it is
   * the session's own backing file / last output, or overwrite is true — a
   * pre-existing import sibling is guarded like any save-as.
   */
  async save(
    rawPath?: string,
    format: 'xlsx' | 'origin' = 'xlsx',
    options: { overwrite?: boolean } = {},
  ): Promise<XlsxSaveResult> {
    // a pinned root that vanished (moved/renamed workspace directory) must
    // fail here, before confinement lets the gateway write anywhere (BUG-1103)
    await assertWorkspaceRootExists(this.root)
    if (format === 'origin') return this.saveToOrigin()
    const target = resolveConfined(rawPath ?? this.defaultTarget(), this.root)
    await assertSaveTargetFree(target, [this.backingPath, ...this.savedTargets], options.overwrite)
    if (target === this.backingPath) await this.assertBackingUnchanged()

    // a fresh (guarded) target promotes exclusively: a file created between
    // the guard's stat and the gateway's write surfaces the clobber error
    // instead of being silently replaced; targets this session owns (or an
    // overwrite) replace by intent
    const replacement =
      options.overwrite === true || target === this.backingPath || this.savedTargets.has(target)
    // BUG-1761: without this overlay an edited save keeps the file's own (or
    // absent) formula caches — script readers (openpyxl data_only, pandas)
    // silently read stale or missing numbers. The sidecar's recalc engine
    // evaluates the journal edits and the results land in <v> via the same
    // gateway overlay the sheets app saves through.
    const { values: formulaValues, note: recalcNote } =
      this.edits.length > 0
        ? await this.recalculateFormulaCaches()
        : { values: [] as SheetFormulaValues[], note: null as string | null }
    let result: Awaited<ReturnType<typeof saveWorkbookViaSidecar>>
    try {
      result = await saveWorkbookViaSidecar({
        client: this.io,
        sourcePath: this.backingPath,
        targetPath: target,
        edits: this.edits.map((edit) => ({ ...edit, cell: { ...edit.cell } })),
        formulaValues,
        exclusiveTarget: !replacement,
      })
    } catch (e) {
      if (e instanceof SaveTargetExistsError) throw saveTargetExistsError(target)
      // UX-1690: a read-only workspace fails inside the sidecar/gateway write
      // and surfaces as a raw errno or the sidecar's errno text; name the
      // actual cause instead (non-permission errors pass through unchanged)
      throw asWorkspaceWriteError(e, target)
    }
    const bytes = await statOrNull(target)
    const unchanged = this.edits.length === 0
    this.edits.length = 0
    this.savedPath = target
    this.savedTargets.add(target)
    if (target === this.backingPath) {
      this.baseline = await statOrNull(this.backingPath)
      // PERF-1778: the in-place save replaced the backing bytes with content
      // the recalc engine's resident model already reflects (the overlay
      // recalculated the same journal). Refresh the resident model's file
      // stamp before the reopen below re-indexes, so the next save's overlay
      // is a resident hit (~ms) instead of a whole-book IronCalc re-import
      // (~20s on a 100k book, previously paid after every single save).
      // Degraded overlays (recalcNote !== null) skip the restamp: the model
      // may lack the saved edits, and only a rebuild from the file is safe.
      // Best-effort by contract: a sidecar predating the command (or any
      // transient error) keeps today's rebuild behavior.
      if (recalcNote === null) {
        try {
          await this.io.restampRecalc(this.backingPath)
        } catch {
          // stamp stays stale; the next recalc rebuilds (pre-existing behavior)
        }
      }
      // the sidecar's in-memory index still reflects the pre-save file:
      // reopen the session so subsequent reads see the saved cells
      await this.reopenSidecar()
    }
    return {
      path: target,
      bytes: bytes?.size ?? 0,
      format: 'xlsx',
      unchanged,
      warnings: [
        ...(this.originPath !== null
          ? [
              `Saved as .xlsx; the original ${this.format === 'ods' ? '.ods' : '.xls'} file "${this.originPath}" was left untouched.`,
            ]
          : []),
        ...(recalcNote !== null
          ? [
              `Formula caches were not refreshed (${recalcNote}). The file keeps the ` +
                'fullCalcOnLoad recalc flag, so spreadsheet apps recalculate on open, but script ' +
                'readers (openpyxl data_only, pandas) see no cached values for the formula cells.',
            ]
          : []),
      ],
      touchedEntries: [...result.touchedEntries],
    }
  }

  /**
   * Fresh cached values for every formula cell of the workbook, evaluated with
   * the pending journal applied (BUG-1761). Best-effort by design: any engine,
   * index or budget failure degrades to an empty overlay plus a reason for the
   * save warning, never to a failed save. The sidecar's lazy worksheet index
   * is waited out (BUG-1776), so the refresh works in the cold flow — a save
   * with no prior read_range in the session — within the same budget.
   *
   * Only numeric results are overlaid: the recalc wire reports text results
   * only through their formatted display string, and baking number formats
   * into a cached <v> would corrupt the value (the fullCalcOnLoad flag stays
   * the mitigation for those cells). Journaled cells are excluded — the
   * journal wins on its own cells, same as the app's overlay.
   */
  private async recalculateFormulaCaches(): Promise<{
    values: SheetFormulaValues[]
    note: string | null
  }> {
    const degrade = (note: string): { values: SheetFormulaValues[]; note: string } => ({
      values: [],
      note,
    })
    const backing = await statOrNull(this.backingPath)
    if (backing !== null && backing.size > RECALC_MAX_SOURCE_BYTES) {
      return degrade('the workbook is above the recalculation engine size budget')
    }

    // every formula cell of the workbook, indexed by the sidecar. The first
    // call per sheet starts the sidecar's lazy background indexer, and the
    // wire command never blocks, so poll until the index completes (BUG-1776)
    const formulaCells: Array<{ sheetName: string; row: number; column: number }> = []
    for (const sheet of this.sheets) {
      const cells = await this.formulaCellsOfSheet(sheet)
      if (typeof cells === 'string') return degrade(cells)
      for (const cellRaw of cells) {
        const cell = asRecord(cellRaw, 'formula cell')
        if (typeof cell.row === 'number' && typeof cell.column === 'number') {
          formulaCells.push({ sheetName: sheet.name, row: cell.row, column: cell.column })
        }
      }
    }
    if (formulaCells.length === 0) return { values: [], note: null }

    // the journal's content edits, as user input for the engine (a journaled
    // formula goes in verbatim); cells the journal writes are excluded from
    // the overlay below — the journal wins on its own cells
    const contentEdits = this.edits.filter((edit) => edit.writeValue)
    if (contentEdits.length > MAX_RECALC_EDITS) {
      return degrade('too many pending edits for the recalculation engine budget')
    }
    const edits = contentEdits.map(
      (
        edit,
      ): {
        sheet: string
        row: number
        column: number
        input: string
      } => {
        const value = edit.cell.value
        const input =
          edit.cell.formula !== undefined
            ? edit.cell.formula
            : value === null
              ? ''
              : typeof value === 'boolean'
                ? value
                  ? 'TRUE'
                  : 'FALSE'
                : String(value)
        return { sheet: edit.sheetName, row: edit.row, column: edit.column, input }
      },
    )
    const journaled = new Set(
      contentEdits.map((edit) => `${edit.sheetName}!${String(edit.row)},${String(edit.column)}`),
    )

    // 1x1 reads over the formula cells, chunked under the engine's read budget
    type OverlayCell = { row: number; column: number; value: number }
    const overlay = new Map<string, OverlayCell[]>()
    for (let start = 0; start < formulaCells.length; start += MAX_RECALC_READ_CELLS) {
      const chunk = formulaCells.slice(start, start + MAX_RECALC_READ_CELLS)
      let raw: unknown
      try {
        raw = await this.io.recalcCells({
          path: this.backingPath,
          edits,
          reads: chunk.map((cell) => ({
            sheet: cell.sheetName,
            range: {
              startRow: cell.row,
              endRow: cell.row,
              startColumn: cell.column,
              endColumn: cell.column,
            },
          })),
        })
      } catch {
        return degrade('the formula engine was unavailable')
      }
      const result = asRecord(raw, 'recalc result')
      if (!Array.isArray(result.cells)) return degrade('the formula engine reply was malformed')
      for (const cellRaw of result.cells) {
        const cell = asRecord(cellRaw, 'recalc cell')
        if (
          cell.isFormula !== true ||
          typeof cell.row !== 'number' ||
          typeof cell.column !== 'number' ||
          typeof cell.number !== 'number' ||
          !Number.isFinite(cell.number)
        ) {
          continue
        }
        const sheetName = typeof cell.sheet === 'string' ? cell.sheet : ''
        if (!this.sheets.some((sheet) => sheet.name === sheetName)) continue
        const key = `${sheetName}!${String(cell.row)},${String(cell.column)}`
        if (journaled.has(key)) continue
        const cells = overlay.get(sheetName) ?? []
        cells.push({ row: cell.row, column: cell.column, value: cell.number })
        overlay.set(sheetName, cells)
      }
    }
    return {
      values: [...overlay].map(([sheetName, cells]) => ({ sheetName, cells })),
      note: null,
    }
  }

  /**
   * The formula cells of one sheet, waiting out the sidecar's lazy worksheet
   * index (BUG-1776). The wire command returns immediately — whatever has been
   * indexed so far plus an `indexingComplete` flag — because the sheets app
   * polls it from its UI, and a blocking wait would stall the sidecar's serial
   * request loop. The overlay needs the finished index, so it polls (the first
   * call also spawns the indexer) and degrades with a named reason when the
   * index errors, reports truncation or misses the wait budget. Returns the
   * cell list, or the degradation note when the index is not usable.
   */
  private async formulaCellsOfSheet(sheet: XlsxSheetSummary): Promise<readonly unknown[] | string> {
    const deadline = Date.now() + recalcIndexWaitMs
    for (;;) {
      let raw: unknown
      try {
        raw = await this.io.readFormulaCells({ sessionId: this.sessionId, sheetId: sheet.id })
      } catch {
        return 'the formula-cell index was unavailable'
      }
      const result = asRecord(raw, 'formula-cells result')
      if (result.truncated === true || !Array.isArray(result.cells)) {
        return 'the formula-cell index was truncated or malformed'
      }
      if (result.indexingComplete === true) return result.cells
      if (Date.now() >= deadline) return 'the sheet index did not finish building in time'
      await delay(RECALC_INDEX_POLL_MS)
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
      // 3. atomic promote onto the original .ods (local capture: the null
      // guard at the top of saveToOrigin does not reach inside the closure)
      const originPath = this.originPath
      const tmpTarget = join(dirname(originPath), `.${basename(originPath)}.airy-${randomUUID()}`)
      try {
        await withSaveTmpCleanup(tmpTarget, async () => {
          await copyFile(output, tmpTarget)
          await rename(tmpTarget, originPath)
        })
      } catch (e) {
        // UX-1690: same friendly mapping as the workbook save — the origin
        // export fails the same way in a read-only workspace
        throw asWorkspaceWriteError(e, originPath)
      }
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

  /**
   * Refresh the sidecar's read index after an in-place save. The fresh
   * session opens BEFORE the stale one closes (PERF-1778): the sidecar
   * releases the resident recalc model only with the last session on the
   * path, so the refresh keeps the model the restamp above just renewed —
   * closing first used to drop it and force every save to pay a cold
   * re-import. Failure ordering improves too: a failed open now leaves the
   * old (stale-index) session alive instead of a closed one.
   */
  private async reopenSidecar(): Promise<void> {
    const previous = this.sessionId
    const openInfo = parseOpenResult(await this.io.open(this.backingPath))
    this.sessionId = openInfo.sessionId
    this.sheets = openInfo.sheets
    this.activeSheetIndex = openInfo.activeTab
    try {
      await this.io.close(previous)
    } catch {
      // a stale session leak is preferable to failing a completed save
    }
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

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms))
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

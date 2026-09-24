// Shared line-session core (REFR-001): the machinery the markdown (PAR-003)
// and html (PAR-004) sessions were independently duplicating after their
// parallel births, minus the parts that genuinely differ. One open plain-text
// file -> a line-oriented in-memory model -> a byte-preserving save: the file
// is split into lines that each keep their OWN terminator, edits splice that
// array, and untouched lines (terminators included) round-trip
// byte-identically on save; a zero-edit save writes the original bytes back
// verbatim, so a file the agent only read never changes on disk.
//
// This module owns: the line model (split/join/EOL dominance, BOM flag kept
// out of the editable text; the phantom entry after a final newline is
// neither counted nor addressable — BUG-1102), the open gauntlet (stat-first
// byte caps, NUL refusal, pre-allocation line cap), the line-ops engine
// (insertLines/replaceLines/deleteLines/findReplace with atomic batch
// validation), the insert-position resolver, the read-selection skeleton, and
// the atomic save with the docx session's fences (drift refusal, save-target
// ownership, tmp + rename, stale-root refusal when the pinned workspace root
// was renamed away — BUG-1103). The sessions stay owners of their true
// divergence points, declared as LineSessionHooks: the charset decode policy,
// the structure summary the read renders, and a couple of kind-specific error
// phrasings.
import { randomUUID } from 'node:crypto'
import type { Stats } from 'node:fs'
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises'
import { dirname, basename, join } from 'node:path'

import {
  asWorkspaceWriteError,
  assertSaveTargetFree,
  FencingError,
  promoteNewFileExclusively,
  withSaveTmpCleanup,
} from '../docx/session.js'
import { assertWorkspaceRootExists, resolveConfined, workspaceRoot } from '../docx/paths.js'
import { replaceCaseInsensitive } from '../case-fold.js'

// ---- limits (mirror the docx session, scaled to the MCP 30k answer budget) ----

const MAX_OPEN_BYTES = 8 * 1024 * 1024
/**
 * Largest line count an open materializes: the byte cap alone does not bound
 * the model — a file of bare EOLs (8 MiB of `\r\n` is ~4M lines) would
 * allocate millions of line objects and hundreds of MB per session. The count
 * runs over the decoded text BEFORE splitLines allocates anything.
 */
const MAX_OPEN_LINES = 2_000_000
const READ_MAX_CHARS = 30_000
/** insert/ops payload caps the sessions quote in their kind-specific errors */
export const INSERT_MAX_CHARS = 200_000
const OPS_TEXT_MAX_CHARS = 200_000
/** how many structure entries (headings) a read summary lists before eliding */
export const HEADING_LIST_MAX = 200
/**
 * Largest range span a read materializes: the request schema does not bound
 * `end`, so the session must reject a huge span BEFORE building the index
 * array (a range like 0..2^53 would otherwise hang/OOM the server).
 */
const RANGE_MAX_SPAN = 10_000

/** U+FEFF as pure-ASCII source (a literal BOM char in source trips tooling) */
const BOM_CHAR = String.fromCharCode(0xfeff)
const NUL_CHAR = String.fromCharCode(0)

/** one addressable line: its text (no terminator) and the terminator it had */
export interface Line {
  text: string
  /** '\n' | '\r\n' | '\r' for terminated lines, '' for the final unterminated one */
  eol: string
}

/** Successful decode result shared by both charset policies */
export interface DecodedText {
  text: string
  /** the bytes started with a byte order mark (re-applied on every save) */
  bom: boolean
  encoding: string
}

export interface LineReadOptions {
  /** line indexes to return in full; must exist (read_document blocks) */
  lines?: number[]
  /** inclusive line range to return in full */
  range?: { start: number; end: number }
}

export interface LineInsertResult {
  /** lines inserted */
  inserted: number
  /** line index the block landed after (-1 = document start) */
  at: number
  lineCount: number
  dirty: boolean
  /** human-readable position summary */
  detail: string
}

export interface LineOpResult {
  op: string
  matched: number
  changed: number
  detail?: string
}

export interface LineSaveResult {
  path: string
  bytes: number
  unchanged: boolean
  warnings: string[]
}

interface FileStamp {
  mtimeMs: number
  size: number
}

/**
 * The sessions' divergence points, injected once at open. Everything else
 * (line model, open gauntlet, ops engine, save fences) is shared behavior
 * that must stay byte-identical across the formats.
 */
export interface LineSessionHooks {
  /** lowercase format label used in cap/refusal/unknown-op error messages */
  kind: 'markdown' | 'html'
  /**
   * Decode the opened bytes (the charset policy is the real divergence:
   * markdown refuses anything past BOM-prefixed UTF-8/UTF-16, html also
   * honors a declared legacy charset). Returns null (not an exception) so
   * open() can phrase one actionable error naming the file.
   */
  decode(bytes: Uint8Array): DecodedText | null
  /** error message when decode returns null */
  refusalMessage(path: string): string
  /**
   * Validate the insertLines op text payload up to (but excluding) the shared
   * length cap: the sessions phrase the required/non-empty errors differently
   * ("text is required ... must be non-empty" vs "text is required (non-empty
   * HTML/markup source)"). Calls `fail` (which throws) and never returns on a
   * bad payload.
   */
  insertOpText(value: unknown, fail: (message: string) => never): string
  /**
   * Runs before an edited save re-encodes (html pins a legacy charset
   * declaration to utf-8 so the re-encoded file still renders; markdown has
   * no hook). May splice the lines in place and push save warnings.
   */
  beforeEncode?(lines: Line[], warnings: string[], encoding: string): void
  /**
   * Marker lines whose content must land BEFORE the line, not after (html:
   * a bare closing </body>/</html>/</head> tag — inserting after it would
   * push the fragment outside the element it closes, DOC-1506). Receives the
   * full text of the first line matching the marker.
   */
  insertBeforeMarkerLine?(lineText: string): boolean
}

// ---- text <-> line model ----

/** Split text into lines, each remembering its own terminator. */
export function splitLines(text: string): Line[] {
  const lines: Line[] = []
  let start = 0
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    if (ch === '\n' || ch === '\r') {
      const eol = ch === '\r' && text[i + 1] === '\n' ? '\r\n' : ch
      lines.push({ text: text.slice(start, i), eol })
      i += eol.length
      start = i
    } else {
      i += 1
    }
  }
  lines.push({ text: text.slice(start), eol: '' })
  return lines
}

/**
 * The line count splitLines would produce, counted WITHOUT materializing the
 * array (the same walk minus the pushes) so an oversize file can be refused
 * before the model allocation.
 */
function countLines(text: string): number {
  let count = 1
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    if (ch === '\n' || ch === '\r') {
      const eolLen = ch === '\r' && text[i + 1] === '\n' ? 2 : 1
      count += 1
      i += eolLen
    } else {
      i += 1
    }
  }
  return count
}

/**
 * The terminator inserted/edited lines get: CRLF when the file uses any,
 * else CR when the file has bare carriage returns, else LF (also for a fresh
 * empty file).
 */
export function dominantEol(lines: readonly Line[]): '\n' | '\r\n' | '\r' {
  let crlf = false
  let bareCr = false
  for (const { eol } of lines) {
    if (eol === '\r\n') crlf = true
    else if (eol === '\r') bareCr = true
  }
  if (crlf) return '\r\n'
  if (bareCr) return '\r'
  return '\n'
}

/**
 * The line count agents see and address. When the file ends WITH a line
 * terminator, splitLines' final entry is a phantom — the zero-width position
 * after the last newline — and it must neither count nor accept indexes
 * (BUG-1102): "a\nb\n" is 2 lines, not 3, and appending at its end must not
 * materialize the phantom as a blank line. A lone unterminated empty entry is
 * the empty file itself and stays addressable, so inserts into it have a
 * position (the empty-file append lands before it, keeping it as the new
 * trailing phantom).
 */
export function addressableLineCount(lines: readonly Line[]): number {
  if (lines.length < 2) return lines.length
  const last = lines[lines.length - 1]!
  return last.text === '' && last.eol === '' ? lines.length - 1 : lines.length
}

/** Join the line model back into text; `bom` re-prepends the leading BOM char. */
function joinLines(lines: readonly Line[], bom: boolean): string {
  return (bom ? BOM_CHAR : '') + lines.map((l) => l.text + l.eol).join('')
}

/**
 * Agent text -> session lines. `lastEol` is what the final line of the block
 * carries: the dominant EOL for a spliced-in block (following content stays on
 * its own line), or '' when the block replaces lines at a file end that had no
 * trailing newline (preserving the file's no-final-newline shape).
 */
function toLines(text: string, eol: string, lastEol: string): Line[] {
  const parts = text.split(/\r\n|\r|\n/)
  return parts.map((t, i) => ({ text: t, eol: i === parts.length - 1 ? lastEol : eol }))
}

/**
 * Give the line at `after` a terminator when it is the file's unterminated
 * last line and content is being spliced in below it (otherwise the first
 * inserted line would merge into it on disk). Returns a note when that
 * changed the file's shape.
 */
function ensureTerminatedBefore(lines: Line[], after: number, eol: string): string | null {
  const target = lines[after]
  if (target && target.eol === '' && after === lines.length - 1) {
    target.eol = eol
    return 'the previous last line had no line break and gained one'
  }
  return null
}

// ---- decode primitives (BOM-aware strict UTF-8/UTF-16, no mojibake fallback) ----

export function decodeStrict(bytes: Uint8Array, charset: string): string | null {
  try {
    return new TextDecoder(charset, { fatal: true, ignoreBOM: true }).decode(bytes)
  } catch {
    return null
  }
}

/** What sniffing the leading BOM found: the shared head of both decode policies. */
export type BomSniff =
  | { status: 'none' }
  | { status: 'invalid' }
  | { status: 'decoded'; text: string; encoding: 'utf-8' | 'utf-16le' | 'utf-16be' }

/**
 * The BOM-prefixed charsets both sessions accept natively: the mark picks its
 * charset and survives into the `bom` flag. `none` hands control back to the
 * session's BOM-less policy; `invalid` means a BOM was present but the bytes
 * did not decode (refuse rather than bless mojibake).
 */
export function sniffBom(bytes: Uint8Array): BomSniff {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    const text = decodeStrict(bytes, 'utf-8')
    return text === null ? { status: 'invalid' } : { status: 'decoded', text, encoding: 'utf-8' }
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    const text = decodeStrict(bytes, 'utf-16le')
    return text === null ? { status: 'invalid' } : { status: 'decoded', text, encoding: 'utf-16le' }
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const text = decodeStrict(bytes, 'utf-16be')
    return text === null ? { status: 'invalid' } : { status: 'decoded', text, encoding: 'utf-16be' }
  }
  return { status: 'none' }
}

// ---- read skeleton ----

/** clip long read output at the MCP answer budget, with an actionable note */
function clip(text: string, max: number, hint: string): string {
  return text.length <= max
    ? text
    : `${text.slice(0, max)}\n...(output truncated at ${String(max)} characters; ${hint})`
}

/** Resolve lines/range-style selections to validated, sorted line indexes. */
function selectedIndexes(options: LineReadOptions, count: number): number[] | null {
  if (options.lines !== undefined) {
    const valid = options.lines.filter((i) => Number.isInteger(i) && i >= 0 && i < count)
    const invalid = options.lines.length - valid.length
    if (invalid > 0) {
      throw new Error(
        `${String(invalid)} of the requested line indexes are out of range (file has ${String(count)} lines)`,
      )
    }
    if (valid.length === 0) throw new Error('No lines selected (empty lines/range)')
    return [...new Set(valid)].sort((a, b) => a - b)
  }
  if (options.range !== undefined) {
    // validate arithmetically and only then materialize: a huge `end` must
    // fail fast instead of allocating the index array first
    const { start, end } = options.range
    if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start) {
      throw new Error(
        `range start/end must be integers with 0 <= start <= end (got start=${start}, end=${end})`,
      )
    }
    const span = end - start + 1
    if (span > RANGE_MAX_SPAN) {
      throw new Error(
        `range ${String(start)}..${String(end)} spans ${String(span)} lines; the cap is ${String(
          RANGE_MAX_SPAN,
        )} per read (split large ranges into smaller reads)`,
      )
    }
    const validEnd = Math.min(end, count - 1)
    const validCount = Math.max(0, validEnd - start + 1)
    const invalid = span - validCount
    if (invalid > 0) {
      throw new Error(
        `${String(invalid)} of the requested line indexes are out of range (file has ${String(count)} lines)`,
      )
    }
    if (validCount === 0) throw new Error('No lines selected (empty lines/range)')
    const indexes: number[] = []
    for (let i = start; i <= validEnd; i++) indexes.push(i)
    return indexes
  }
  return null
}

/**
 * The read tail both sessions share: after their kind-specific header and
 * structure block, the full text (EOLs normalized to LF), or exactly the
 * selected lines, all under the shared 30k clip budget.
 */
export function renderRead(
  header: string,
  structureBlock: string,
  lines: readonly Line[],
  options: LineReadOptions,
): string {
  // addressable count: the trailing phantom (final newline) is neither
  // selectable nor announced (BUG-1102)
  const selected = selectedIndexes(options, addressableLineCount(lines))
  if (selected === null) {
    const fullText = lines.map((l) => l.text).join('\n')
    return clip(
      [header, structureBlock, '', 'Full text (EOLs normalized to LF):', fullText].join('\n'),
      READ_MAX_CHARS,
      'request a line range to read the rest',
    )
  }
  const body = selected.map((i) => lines[i]!.text).join('\n')
  return clip(
    [
      header,
      structureBlock,
      '',
      `Selected ${String(selected.length)} line(s) (EOLs normalized to LF):`,
      body,
    ].join('\n'),
    READ_MAX_CHARS,
    'request a narrower selection',
  )
}

// ---- the document ----

/**
 * The stateful half of a line session: opened bytes, the line model, the
 * edit journal (dirty), and the save fences. The MarkdownSession and
 * HtmlSession classes wrap one of these and layer their structure model,
 * charset policy, and read formatting on top.
 */
export class LineDocument {
  readonly path: string
  readonly bom: boolean
  readonly encoding: string

  /** confinement root captured at open (save must not follow a later drift) */
  private readonly root: string
  private readonly originalBytes: Uint8Array
  private lineModel: Line[]
  private baseline: FileStamp | null
  private readonly savedTargets = new Set<string>()
  private dirty = false
  private readonly hooks: LineSessionHooks

  private constructor(
    path: string,
    root: string,
    originalBytes: Uint8Array,
    lines: Line[],
    bom: boolean,
    encoding: string,
    stamp: FileStamp | null,
    hooks: LineSessionHooks,
  ) {
    this.path = path
    this.root = root
    this.originalBytes = originalBytes
    this.lineModel = lines
    this.bom = bom
    this.encoding = encoding
    this.baseline = stamp
    this.hooks = hooks
  }

  /** Open a plain-text file inside the workspace root (the shared gauntlet). */
  static async open(
    rawPath: string,
    root: string | undefined,
    hooks: LineSessionHooks,
  ): Promise<{ handle: string; doc: LineDocument }> {
    const path = resolveConfined(rawPath, root)
    // stat first: an oversize file is refused by its size BEFORE the whole
    // content is read into memory (the post-read check stays as a backstop
    // against the file growing between stat and read)
    let info: Stats
    try {
      info = await stat(path)
    } catch (e) {
      throw new Error(`Cannot read "${path}": ${e instanceof Error ? e.message : String(e)}`, {
        cause: e,
      })
    }
    if (info.size > MAX_OPEN_BYTES) {
      throw new Error(
        `"${path}" is ${String(info.size)} bytes; ${hooks.kind} sessions cap at ` +
          `${String(MAX_OPEN_BYTES)} bytes (8 MiB).`,
      )
    }
    let bytes: Uint8Array
    try {
      bytes = new Uint8Array(await readFile(path))
    } catch (e) {
      throw new Error(`Cannot read "${path}": ${e instanceof Error ? e.message : String(e)}`, {
        cause: e,
      })
    }
    const stamp: FileStamp = { mtimeMs: info.mtimeMs, size: info.size }
    if (bytes.byteLength > MAX_OPEN_BYTES) {
      throw new Error(
        `"${path}" is ${String(bytes.byteLength)} bytes; ${hooks.kind} sessions cap at ` +
          `${String(MAX_OPEN_BYTES)} bytes (8 MiB).`,
      )
    }
    const decoded = hooks.decode(bytes)
    if (decoded === null) {
      throw new Error(hooks.refusalMessage(path))
    }
    if (decoded.text.includes(NUL_CHAR)) {
      throw new Error(
        `Cannot open "${path}": the file contains NUL bytes - it does not look like a text document.`,
      )
    }
    // the decoder keeps the BOM as a leading U+FEFF; it lives in the flag and
    // is re-applied on save, so it must not stay in the editable text
    const text =
      decoded.bom && decoded.text.startsWith(BOM_CHAR) ? decoded.text.slice(1) : decoded.text
    // the line count is checked over the decoded text BEFORE splitLines
    // allocates one object per line (the byte cap alone does not bound the
    // model: 8 MiB of bare EOLs is ~4M line objects, hundreds of MB)
    const lineCount = countLines(text)
    if (lineCount > MAX_OPEN_LINES) {
      throw new Error(
        `"${path}" has ${String(lineCount)} lines; ${hooks.kind} sessions cap at ${String(
          MAX_OPEN_LINES,
        )} lines (a file of bare line breaks would otherwise materialize millions of ` +
          'line objects). Split the file and retry.',
      )
    }
    return {
      handle: randomUUID(),
      doc: new LineDocument(
        path,
        root ?? workspaceRoot(),
        bytes,
        splitLines(text),
        decoded.bom,
        decoded.encoding,
        stamp,
        hooks,
      ),
    }
  }

  /** the line model, read-only for the wrapping session's scans and reads */
  get lines(): readonly Line[] {
    return this.lineModel
  }

  get lineCount(): number {
    // the phantom after a final newline is not a line (BUG-1102)
    return addressableLineCount(this.lineModel)
  }

  get isDirty(): boolean {
    return this.dirty
  }

  /** the editable text joined back (no BOM; word/char counts run over this) */
  get text(): string {
    return joinLines(this.lineModel, false)
  }

  // ---- editing ----

  /**
   * Insert text as whole lines: after a marker line (first exact substring
   * match), after heading N (markdown only - the caller resolves the ordinal
   * to a line first), after line `at` (-1 = document start), or at the end by
   * default. Precedence: marker > afterHeading > at. `detailSuffix` lets the
   * html session append its "verbatim" note.
   */
  insert(
    text: string,
    position: { at?: number; marker?: string },
    options: {
      afterHeading?: { ordinal: number; line: number }
      detailSuffix?: string
    } = {},
  ): LineInsertResult {
    // addressable count: a phantom after the final newline is not a position
    // an insert can name (BUG-1102) — "the end" is the last real line, and
    // the splice lands between it and the phantom
    const count = addressableLineCount(this.lineModel)
    const eol = dominantEol(this.lineModel)
    let after: number
    let where: string
    if (position.marker !== undefined) {
      if (position.marker.length === 0) throw new Error('marker must be a non-empty string')
      const found = this.lineModel.findIndex((l) => l.text.includes(position.marker!))
      if (found === -1) {
        throw new Error(
          `marker "${position.marker}" does not match any line - re-read the document and retry`,
        )
      }
      // DOC-1506: a marker that names a structural closing tag (e.g.
      // "</body>") must not push content PAST the element it closes - the
      // session hook opts such lines into a before-the-line insert
      if (this.hooks.insertBeforeMarkerLine?.(this.lineModel[found]!.text) === true) {
        after = found - 1
        where = `before line ${String(found)}, the first line matching marker "${position.marker}"`
      } else {
        after = found
        where = `after marker line ${String(found)}`
      }
    } else if (options.afterHeading !== undefined) {
      after = options.afterHeading.line
      where = `after heading ${String(options.afterHeading.ordinal)} (line ${String(after)})`
    } else if (position.at !== undefined) {
      const index = Math.trunc(position.at)
      if (index < -1 || index > count - 1) {
        throw new Error(
          `at ${String(index)} is out of range (line indexes are 0..${String(count - 1)}; ` +
            'use -1 for the document start) - re-read the document after edits, indexes shift',
        )
      }
      after = index
      where = after === -1 ? 'at the document start' : `after line ${String(after)}`
    } else {
      after = count - 1
      where = 'at the end of the document'
    }
    // the empty file's lone empty line is the whole file, not a phantom to
    // terminate: the block lands at the start and the lone line stays as the
    // trailing phantom — no leading blank line (BUG-1102)
    const loneEmptyLine =
      after === 0 &&
      this.lineModel.length === 1 &&
      this.lineModel[0]!.text === '' &&
      this.lineModel[0]!.eol === ''
    const shapeNote = loneEmptyLine ? null : ensureTerminatedBefore(this.lineModel, after, eol)
    // appending at the end preserves the file's trailing-newline shape: a
    // file that ended without a newline keeps ending without one (shapeNote
    // firing means exactly that case); a file that ended WITH one keeps it —
    // the block is terminated and lands before the trailing phantom
    const newLines = toLines(text, eol, loneEmptyLine || shapeNote === null ? eol : '')
    this.lineModel.splice(loneEmptyLine ? after : after + 1, 0, ...newLines)
    this.dirty = true
    const detail = `inserted ${String(newLines.length)} line(s) ${where}${
      shapeNote ? ` (${shapeNote})` : ''
    }${options.detailSuffix ?? ''}`
    return {
      inserted: newLines.length,
      at: after,
      lineCount: addressableLineCount(this.lineModel),
      dirty: true,
      detail,
    }
  }

  /** Apply a batch of line ops atomically (validated/applied on a copy). */
  applyOps(
    ops: Array<Record<string, unknown>>,
    dryRun = false,
  ): { results: LineOpResult[]; summary: string; dryRun: boolean } {
    const work = this.lineModel.map((l) => ({ ...l }))
    const results: LineOpResult[] = []
    const fail = (message: string): never => {
      throw new Error(`${message} - nothing was applied (atomic); fix and resend the whole batch`)
    }
    // addressable count: ops cannot name the phantom after a final newline
    // (BUG-1102) — the same count reads and insert_content expose
    const lineCount = () => addressableLineCount(work)
    const checkRange = (name: string, from: unknown, to: unknown) => {
      if (
        !Number.isInteger(from) ||
        !Number.isInteger(to) ||
        (from as number) < 0 ||
        (to as number) < (from as number) ||
        (to as number) >= lineCount()
      ) {
        fail(
          `op ${name}: from/to must satisfy 0 <= from <= to < ${String(lineCount())} ` +
            `(got from=${String(from)}, to=${String(to)}) - re-read the document after edits, ` +
            'line indexes shift',
        )
      }
    }
    const allowed = (op: Record<string, unknown>, name: string, keys: readonly string[]) => {
      for (const key of Object.keys(op)) {
        if (key !== 'op' && !keys.includes(key)) fail(`op ${name}: unknown field "${key}"`)
      }
    }
    for (const op of ops) {
      const name = typeof op.op === 'string' ? op.op : '(missing op name)'
      switch (name) {
        case 'insertLines': {
          allowed(op, name, ['after', 'text'])
          const after = op.after
          if (
            !Number.isInteger(after) ||
            (after as number) < -1 ||
            (after as number) > lineCount() - 1
          ) {
            fail(
              `op insertLines: after must be an integer -1..${String(lineCount() - 1)} ` +
                `(got ${String(after)})`,
            )
          }
          // the hook owns the required/non-empty phrasing (it differs per
          // format); the length cap itself is shared
          const text = this.hooks.insertOpText(op.text, fail)
          if (text.length > OPS_TEXT_MAX_CHARS) {
            fail(
              `op insertLines: text is ${String(text.length)} characters; the cap is ${String(
                OPS_TEXT_MAX_CHARS,
              )}`,
            )
          }
          const eol = dominantEol(work)
          // same empty-file handling as insert_content: the block lands at
          // the start, the lone empty line becomes the trailing phantom —
          // no leading blank line (BUG-1102)
          const loneEmptyLine =
            after === 0 && work.length === 1 && work[0]!.text === '' && work[0]!.eol === ''
          const shapeNote = loneEmptyLine
            ? null
            : ensureTerminatedBefore(work, after as number, eol)
          // same trailing-newline preservation as insert_content
          const inserted = toLines(text, eol, loneEmptyLine || shapeNote === null ? eol : '')
          work.splice((after as number) + (loneEmptyLine ? 0 : 1), 0, ...inserted)
          results.push({
            op: name,
            matched: 1,
            changed: inserted.length,
            detail: `inserted ${String(inserted.length)} line(s) after line ${String(after)}${
              shapeNote ? ` (${shapeNote})` : ''
            }`,
          })
          break
        }
        case 'replaceLines': {
          allowed(op, name, ['from', 'to', 'text'])
          checkRange(name, op.from, op.to)
          const text =
            op.text === undefined || op.text === null
              ? undefined
              : typeof op.text === 'string'
                ? op.text
                : fail('op replaceLines: text must be a string')
          if (text !== undefined && text.length > OPS_TEXT_MAX_CHARS) {
            fail(
              `op replaceLines: text is ${String(text.length)} characters; the cap is ${String(
                OPS_TEXT_MAX_CHARS,
              )}`,
            )
          }
          const from = op.from as number
          const to = op.to as number
          const eol = dominantEol(work)
          // replacing the file's unterminated last line keeps that shape
          const lastEol = to === lineCount() - 1 && work[to]!.eol === '' ? '' : eol
          const replacement = text === undefined || text === '' ? [] : toLines(text, eol, lastEol)
          work.splice(from, to - from + 1, ...replacement)
          results.push({
            op: name,
            matched: 1,
            changed: replacement.length,
            detail: `replaced lines ${String(from)}-${String(to)} with ${String(
              replacement.length,
            )} line(s)`,
          })
          break
        }
        case 'deleteLines': {
          allowed(op, name, ['from', 'to'])
          checkRange(name, op.from, op.to)
          const from = op.from as number
          const to = op.to as number
          work.splice(from, to - from + 1)
          results.push({
            op: name,
            matched: 1,
            changed: to - from + 1,
            detail: `deleted lines ${String(from)}-${String(to)}`,
          })
          break
        }
        case 'findReplace': {
          allowed(op, name, ['find', 'replace', 'matchCase', 'from', 'to'])
          const find =
            typeof op.find === 'string' && op.find.length > 0
              ? op.find
              : fail('op findReplace: find must be a non-empty string')
          if (/[\r\n]/.test(find)) {
            fail('op findReplace: find must not contain line breaks (replace is line-scoped)')
          }
          const replace =
            typeof op.replace === 'string'
              ? op.replace
              : fail('op findReplace: replace must be a string')
          // a replace with embedded line breaks would leave an EOL inside one
          // line object, breaking the line-model invariant (every op and the
          // dominant-EOL accounting assume one line = no inner EOL)
          if (/[\r\n]/.test(replace)) {
            fail(
              'op findReplace: replace must not contain line breaks ' +
                '(use insertLines or replaceLines for multi-line edits)',
            )
          }
          const matchCase = op.matchCase !== false
          const from = op.from === undefined ? 0 : op.from
          const to = op.to === undefined ? lineCount() - 1 : op.to
          checkRange(name, from, to)
          let occurrences = 0
          let changedLines = 0
          for (let i = from as number; i <= (to as number); i++) {
            const hay = work[i]!
            if (matchCase) {
              if (!hay.text.includes(find)) continue
              occurrences += hay.text.split(find).length - 1
              work[i] = { ...hay, text: hay.text.split(find).join(replace) }
            } else {
              // the shared fold-safe replace: lowered indices cannot slice
              // the original, where İ (U+0130) shifts every position after
              // it by expanding to two code units (BUG-1101)
              const outcome = replaceCaseInsensitive(hay.text, find, replace)
              if (outcome.count === 0) continue
              occurrences += outcome.count
              work[i] = { ...hay, text: outcome.text }
            }
            changedLines += 1
          }
          results.push({
            op: name,
            matched: occurrences,
            changed: changedLines,
            ...(occurrences === 0 ? { detail: 'no matches (nothing changed)' } : {}),
          })
          break
        }
        default:
          fail(
            `unknown op "${name}" - ${this.hooks.kind} sessions accept insertLines, replaceLines, ` +
              'deleteLines, findReplace',
          )
      }
    }
    if (!dryRun) {
      this.lineModel = work
      this.dirty = true
    }
    return {
      results,
      summary: results
        .map(
          (r) =>
            `${r.op}: matched ${r.matched}, changed ${r.changed}${r.detail ? ` (${r.detail})` : ''}`,
        )
        .join('; '),
      dryRun,
    }
  }

  // ---- saving ----

  /**
   * Save atomically (tmp + rename) with the docx session's fences: saving over
   * the opened file refuses when it changed on disk since open; a target that
   * exists is refused unless the session owns it or overwrite is true. With no
   * edits the original bytes round-trip verbatim (an untouched file never
   * changes on disk, whatever its encoding was); an edited save writes UTF-8
   * with the original BOM flag re-applied (the html session additionally pins
   * a legacy charset declaration via its beforeEncode hook).
   */
  async save(rawPath?: string, options: { overwrite?: boolean } = {}): Promise<LineSaveResult> {
    // a pinned root that vanished (moved/renamed workspace directory) must
    // fail here, before confinement + mkdir silently resurrect it (BUG-1103)
    await assertWorkspaceRootExists(this.root)
    // the root captured at open, not the live one: a drifted
    // AIRY_WORKSPACE_ROOT/cwd between open and save must not re-confine the
    // session (the docx session has the same pinned-root semantics)
    const target = resolveConfined(rawPath ?? this.path, this.root)
    await assertSaveTargetFree(target, [this.path, ...this.savedTargets], options.overwrite)
    if (target === this.path && this.baseline) {
      let current: FileStamp
      try {
        const info = await stat(this.path)
        current = { mtimeMs: info.mtimeMs, size: info.size }
      } catch {
        current = { mtimeMs: -1, size: -1 }
      }
      if (current.mtimeMs !== this.baseline.mtimeMs || current.size !== this.baseline.size) {
        throw new FencingError(this.path)
      }
    }
    const warnings: string[] = []
    let bytes: Uint8Array
    let unchanged = false
    if (!this.dirty) {
      bytes = this.originalBytes
      unchanged = true
    } else {
      this.hooks.beforeEncode?.(this.lineModel, warnings, this.encoding)
      bytes = new TextEncoder().encode(joinLines(this.lineModel, this.bom))
      if (this.encoding !== 'utf-8') {
        warnings.push(`Original encoding was ${this.encoding}; the edited copy is saved as UTF-8.`)
      }
    }
    await mkdir(dirname(target), { recursive: true })
    // basename, not a '/'-split: on Windows the split leaves the whole path
    // in the temp name and writeFile fails on the colons/backslashes
    const tmp = join(
      dirname(target),
      `.${basename(target) || this.hooks.kind}.airy-${randomUUID()}`,
    )
    try {
      await withSaveTmpCleanup(tmp, async () => {
        await writeFile(tmp, bytes)
        if (options.overwrite === true || target === this.path || this.savedTargets.has(target)) {
          await rename(tmp, target)
        } else {
          await promoteNewFileExclusively(tmp, target)
        }
      })
    } catch (e) {
      // UX-1690: a read-only workspace/target fails the tmp write with a raw
      // errno naming the dot-temp file; name the actual cause instead
      throw asWorkspaceWriteError(e, target)
    }
    if (target === this.path) {
      try {
        const info = await stat(this.path)
        this.baseline = { mtimeMs: info.mtimeMs, size: info.size }
      } catch {
        this.baseline = null
      }
    }
    this.savedTargets.add(target)
    return { path: target, bytes: bytes.byteLength, unchanged, warnings }
  }
}

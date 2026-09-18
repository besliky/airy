// Headless Markdown editing session (PAR-003): one open .md/.markdown file
// -> line-oriented in-memory model -> byte-preserving save. Markdown is plain
// text, so the session model is simpler than the docx twin: the file is split
// into lines that each keep their OWN terminator, edits splice that array, and
// untouched lines (terminators included) round-trip byte-identically on save.
// A zero-edit save writes the original bytes back verbatim, so a file the
// agent only read never changes on disk.
//
// Encoding policy (see CHANGELOG 0.9.3 - BOM/U+FEFF semantics): markdown is
// UTF-8 by spec. The session accepts UTF-8 with or without a BOM (also UTF-16
// with a BOM, which the suite's text decoder reads); anything that is not
// valid UTF-8/UTF-16 is refused with an actionable error instead of silently
// decoding mojibake that a save would then bless. A leading BOM is stored as a
// flag, kept out of the editable text, and written back on every save.
//
// Structure model for read_document: ATX headings (`#`..`######`) outside
// fenced code blocks and YAML front matter, reported with their line index;
// setext headings (===/--- underlines) are deliberately not parsed and are
// documented as such - agents address them as plain lines.
import { randomUUID } from 'node:crypto'
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'

import {
  assertSaveTargetFree,
  countWords,
  FencingError,
  promoteNewFileExclusively,
} from '../docx/session.js'
import { resolveConfined } from '../docx/paths.js'

// ---- limits (mirror the docx session, scaled to the MCP 30k answer budget) ----

const MAX_OPEN_BYTES = 8 * 1024 * 1024
const READ_MAX_CHARS = 30_000
const INSERT_MAX_CHARS = 200_000
const OPS_TEXT_MAX_CHARS = 200_000
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
export interface MarkdownLine {
  text: string
  /** '\n' | '\r\n' | '\r' for terminated lines, '' for the final unterminated one */
  eol: string
}

export interface MarkdownHeading {
  /** 1-based ordinal agents address via insert_content afterHeading */
  ordinal: number
  /** 0-based line index of the `#` line */
  line: number
  level: number
  text: string
}

export interface MarkdownSessionMeta {
  handle: string
  kind: 'markdown'
  path: string
  fileName: string
  format: 'md'
  converted: false
  editable: true
  warnings: string[]
  lineCount: number
  headingCount: number
  wordCount: number
  charCount: number
  /** dominant line ending: what inserted/edited lines get */
  eol: '\n' | '\r\n' | '\r'
  /** the file started with a byte order mark */
  bom: boolean
  dirty: boolean
}

export interface MarkdownReadOptions {
  /** line indexes to return in full; must exist (read_document blocks) */
  lines?: number[]
  /** inclusive line range to return in full */
  range?: { start: number; end: number }
}

export interface MarkdownSaveResult {
  path: string
  bytes: number
  unchanged: boolean
  warnings: string[]
}

export interface InsertResult {
  /** lines inserted */
  inserted: number
  /** line index the block landed after (-1 = document start) */
  at: number
  lineCount: number
  dirty: boolean
  /** human-readable position summary */
  detail: string
}

export interface MarkdownOpResult {
  op: string
  matched: number
  changed: number
  detail?: string
}

interface FileStamp {
  mtimeMs: number
  size: number
}

// ---- text <-> line model ----

/** Split text into lines, each remembering its own terminator. */
export function splitLines(text: string): MarkdownLine[] {
  const lines: MarkdownLine[] = []
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
 * The terminator inserted/edited lines get: CRLF when the file uses any,
 * else CR when the file has bare carriage returns, else LF (also for a fresh
 * empty file).
 */
function dominantEol(lines: readonly MarkdownLine[]): '\n' | '\r\n' | '\r' {
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

/** Join the line model back into text; `bom` re-prepends the leading BOM char. */
function joinLines(lines: readonly MarkdownLine[], bom: boolean): string {
  return (bom ? BOM_CHAR : '') + lines.map((l) => l.text + l.eol).join('')
}

/**
 * Agent text -> session lines. `lastEol` is what the final line of the block
 * carries: the dominant EOL for a spliced-in block (following content stays on
 * its own line), or '' when the block replaces lines at a file end that had no
 * trailing newline (preserving the file's no-final-newline shape).
 */
function toLines(text: string, eol: string, lastEol: string): MarkdownLine[] {
  const parts = text.split(/\r\n|\r|\n/)
  return parts.map((t, i) => ({ text: t, eol: i === parts.length - 1 ? lastEol : eol }))
}

// ---- structure scan ----

const ATX_RE = /^ {0,3}(#{1,6})(?:[ \t]+(.*?))?[ \t]*$/
const FENCE_RE = /^ {0,3}(`{3,}|~{3,})/
const FRONT_MATTER_OPEN_RE = /^---[ \t]*$/
const FRONT_MATTER_CLOSE_RE = /^(?:---|\.{3})[ \t]*$/

/** Strip an ATX closing hash sequence ("## Title ##" -> "Title"). */
function stripClosingHashes(text: string): string {
  const m = /^(.*?)[ \t]+#+[ \t]*$/.exec(text)
  if (m) return m[1]!
  return text.trimEnd()
}

/**
 * ATX headings outside fences and YAML front matter, in document order. The
 * ordinal (1-based) is the number agents pass to insert_content afterHeading.
 */
export function scanHeadings(lines: readonly MarkdownLine[]): MarkdownHeading[] {
  const headings: MarkdownHeading[] = []
  let fence: { char: string; len: number } | null = null
  let inFrontMatter = lines.length > 0 && FRONT_MATTER_OPEN_RE.test(lines[0]!.text)
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i]!.text
    if (inFrontMatter) {
      if (i > 0 && FRONT_MATTER_CLOSE_RE.test(raw)) inFrontMatter = false
      continue
    }
    const fenceMatch = FENCE_RE.exec(raw)
    if (fence) {
      // a closing fence must use the same character and be at least as long
      if (fenceMatch && fenceMatch[1]![0] === fence.char && fenceMatch[1]!.length >= fence.len) {
        fence = null
      }
      continue
    }
    if (fenceMatch) {
      fence = { char: fenceMatch[1]![0]!, len: fenceMatch[1]!.length }
      continue
    }
    const atx = ATX_RE.exec(raw)
    if (atx) {
      const body = atx[2] ?? ''
      headings.push({
        ordinal: headings.length + 1,
        line: i,
        level: atx[1]!.length,
        text: body === '' ? '' : stripClosingHashes(body),
      })
    }
  }
  return headings
}

// ---- decode (BOM-aware strict UTF-8/UTF-16, no mojibake fallback) ----

function decodeStrict(bytes: Uint8Array, charset: string): string | null {
  try {
    return new TextDecoder(charset, { fatal: true, ignoreBOM: true }).decode(bytes)
  } catch {
    return null
  }
}

/**
 * Decode markdown bytes: a BOM picks its charset and survives as a returned
 * flag; BOM-less files must be valid UTF-8. Returns null (not an exception)
 * so open() can phrase one actionable error naming the file.
 */
function decodeMarkdown(
  bytes: Uint8Array,
): { text: string; bom: boolean; encoding: 'utf-8' | 'utf-16le' | 'utf-16be' } | null {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    const text = decodeStrict(bytes, 'utf-8')
    return text === null ? null : { text, bom: true, encoding: 'utf-8' }
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    const text = decodeStrict(bytes, 'utf-16le')
    return text === null ? null : { text, bom: true, encoding: 'utf-16le' }
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const text = decodeStrict(bytes, 'utf-16be')
    return text === null ? null : { text, bom: true, encoding: 'utf-16be' }
  }
  const text = decodeStrict(bytes, 'utf-8')
  return text === null ? null : { text, bom: false, encoding: 'utf-8' }
}

// ---- insert position params ----

export interface InsertPosition {
  /** insert after this line index (-1 = document start) */
  at?: number
  /** insert after heading N (1-based ordinal from the read structure) */
  afterHeading?: number
  /** insert after the first line containing this exact substring */
  marker?: string
}

// ---- session ----

export class MarkdownSession {
  readonly handle: string
  readonly path: string

  private readonly originalBytes: Uint8Array
  private lines: MarkdownLine[]
  private readonly bom: boolean
  private readonly encoding: 'utf-8' | 'utf-16le' | 'utf-16be'
  private baseline: FileStamp | null
  private readonly savedTargets = new Set<string>()
  private dirty = false

  private constructor(
    handle: string,
    path: string,
    originalBytes: Uint8Array,
    lines: MarkdownLine[],
    bom: boolean,
    encoding: 'utf-8' | 'utf-16le' | 'utf-16be',
    stamp: FileStamp | null,
  ) {
    this.handle = handle
    this.path = path
    this.originalBytes = originalBytes
    this.lines = lines
    this.bom = bom
    this.encoding = encoding
    this.baseline = stamp
  }

  /** Open a .md/.markdown file inside the workspace root. */
  static async open(rawPath: string, root?: string): Promise<MarkdownSession> {
    const path = resolveConfined(rawPath, root)
    let bytes: Uint8Array
    let stamp: FileStamp
    try {
      bytes = new Uint8Array(await readFile(path))
      const info = await stat(path)
      stamp = { mtimeMs: info.mtimeMs, size: info.size }
    } catch (e) {
      throw new Error(`Cannot read "${path}": ${e instanceof Error ? e.message : String(e)}`, {
        cause: e,
      })
    }
    if (bytes.byteLength > MAX_OPEN_BYTES) {
      throw new Error(
        `"${path}" is ${String(bytes.byteLength)} bytes; markdown sessions cap at ` +
          `${String(MAX_OPEN_BYTES)} bytes (8 MiB).`,
      )
    }
    const decoded = decodeMarkdown(bytes)
    if (decoded === null) {
      throw new Error(
        `Cannot open "${path}": the bytes are not valid UTF-8 (markdown is a UTF-8 format; ` +
          'BOM-prefixed UTF-8/UTF-16 is accepted). Convert the file and retry.',
      )
    }
    // the decoder keeps the BOM as a leading U+FEFF; it lives in the flag and
    // is re-applied on save, so it must not stay in the editable text
    const text =
      decoded.bom && decoded.text.startsWith(BOM_CHAR) ? decoded.text.slice(1) : decoded.text
    if (decoded.text.includes(NUL_CHAR)) {
      throw new Error(
        `Cannot open "${path}": the file contains NUL bytes - it does not look like a text document.`,
      )
    }
    return new MarkdownSession(
      randomUUID(),
      path,
      bytes,
      splitLines(text),
      decoded.bom,
      decoded.encoding,
      stamp,
    )
  }

  // ---- reading ----

  meta(): MarkdownSessionMeta {
    const text = joinLines(this.lines, false)
    return {
      handle: this.handle,
      kind: 'markdown',
      path: this.path,
      fileName: this.path.split('/').pop() ?? this.path,
      format: 'md',
      converted: false,
      editable: true,
      warnings:
        this.encoding === 'utf-8'
          ? []
          : [`Original encoding is UTF-16 (${this.encoding}); saving an edit writes UTF-8.`],
      lineCount: this.lines.length,
      headingCount: scanHeadings(this.lines).length,
      wordCount: countWords(text),
      charCount: text.length,
      eol: dominantEol(this.lines),
      bom: this.bom,
      dirty: this.dirty,
    }
  }

  private headings(): MarkdownHeading[] {
    return scanHeadings(this.lines)
  }

  /**
   * Agent-facing read: file header (stats + EOL/BOM), the heading list with
   * line positions, then the full text - truncated at the 30k budget with a
   * hint to request line ranges instead. With lines/range selected, the full
   * text section is replaced by exactly those lines.
   */
  readDocument(options: MarkdownReadOptions = {}): string {
    const meta = this.meta()
    const header =
      `"${meta.fileName}" - ${String(meta.lineCount)} lines, ${String(meta.wordCount)} words, ` +
      `${String(meta.charCount)} characters. Line endings: ${
        meta.eol === '\r\n' ? 'CRLF' : meta.eol === '\r' ? 'CR' : 'LF'
      }.${meta.bom ? ' Starts with a BOM.' : ''}`
    const headingLines = this.headings().map(
      (h) => `${h.ordinal}|${h.line}|${'#'.repeat(h.level)}|${h.text}`,
    )
    const headingsBlock = [
      `Headings (ordinal|line|level|text):${headingLines.length === 0 ? ' none' : ''}`,
      ...headingLines,
    ].join('\n')

    const selected = this.selectedIndexes(options)
    if (selected === null) {
      const fullText = this.lines.map((l) => l.text).join('\n')
      const clipped = clip(fullText, READ_MAX_CHARS, 'request a line range to read the rest')
      return [header, headingsBlock, '', 'Full text (EOLs normalized to LF):', clipped].join('\n')
    }
    const body = selected.map((i) => this.lines[i]!.text).join('\n')
    const clipped = clip(body, READ_MAX_CHARS, 'request a narrower selection')
    return [
      header,
      headingsBlock,
      '',
      `Selected ${String(selected.length)} line(s) (EOLs normalized to LF):`,
      clipped,
    ].join('\n')
  }

  /** Resolve blocks/range-style selections to validated, sorted line indexes. */
  private selectedIndexes(options: MarkdownReadOptions): number[] | null {
    const count = this.lines.length
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

  // ---- editing ----

  /**
   * Give the line at `after` a terminator when it is the file's unterminated
   * last line and content is being spliced in below it (otherwise the first
   * inserted line would merge into it on disk). Returns a note when that
   * changed the file's shape.
   */
  private static ensureTerminatedBefore(
    lines: MarkdownLine[],
    after: number,
    eol: string,
  ): string | null {
    const target = lines[after]
    if (target && target.eol === '' && after === lines.length - 1) {
      target.eol = eol
      return 'the previous last line had no line break and gained one'
    }
    return null
  }

  /**
   * Insert markdown text at a position: after a marker line (first exact
   * substring match), after heading N (1-based ordinal from the structure
   * list), after line `at` (-1 = document start), or at the end by default.
   * Precedence: marker > afterHeading > at.
   */
  insertContent(text: string, position: InsertPosition = {}): InsertResult {
    if (typeof text !== 'string' || text.length === 0) {
      throw new Error('text is required for a markdown insert (markdown source, not HTML)')
    }
    if (text.length > INSERT_MAX_CHARS) {
      throw new Error(
        `insert text is ${String(text.length)} characters; the cap is ${String(INSERT_MAX_CHARS)}.`,
      )
    }
    const count = this.lines.length
    const eol = dominantEol(this.lines)
    let after: number
    let where: string
    if (position.marker !== undefined) {
      if (position.marker.length === 0) throw new Error('marker must be a non-empty string')
      const found = this.lines.findIndex((l) => l.text.includes(position.marker!))
      if (found === -1) {
        throw new Error(
          `marker "${position.marker}" does not match any line - re-read the document and retry`,
        )
      }
      after = found
      where = `after marker line ${String(found)}`
    } else if (position.afterHeading !== undefined) {
      const headings = this.headings()
      const ordinal = Math.trunc(position.afterHeading)
      if (ordinal < 1 || ordinal > headings.length) {
        throw new Error(
          `afterHeading ${String(ordinal)} is out of range (the document has ${String(
            headings.length,
          )} heading(s)) - re-read the document for the current ordinals`,
        )
      }
      after = headings[ordinal - 1]!.line
      where = `after heading ${String(ordinal)} (line ${String(after)})`
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
    const shapeNote = MarkdownSession.ensureTerminatedBefore(this.lines, after, eol)
    // appending at the end preserves the file's trailing-newline shape: a
    // file that ended without a newline keeps ending without one (shapeNote
    // firing means exactly that case)
    const newLines = toLines(text, eol, shapeNote !== null ? '' : eol)
    this.lines.splice(after + 1, 0, ...newLines)
    this.dirty = true
    const detail = `inserted ${String(newLines.length)} line(s) ${where}${
      shapeNote ? ` (${shapeNote})` : ''
    }`
    return {
      inserted: newLines.length,
      at: after,
      lineCount: this.lines.length,
      dirty: true,
      detail,
    }
  }

  /** Apply a batch of line ops atomically (validated/applied on a copy). */
  applyOps(
    ops: Array<Record<string, unknown>>,
    dryRun = false,
  ): { results: MarkdownOpResult[]; summary: string; dryRun: boolean } {
    const work = this.lines.map((l) => ({ ...l }))
    const results: MarkdownOpResult[] = []
    const fail = (message: string): never => {
      throw new Error(`${message} - nothing was applied (atomic); fix and resend the whole batch`)
    }
    const lineCount = () => work.length
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
    const takeText = (name: string, value: unknown, required: boolean): string => {
      if (value === undefined || value === null) {
        if (required) fail(`op ${name}: text is required`)
        return ''
      }
      if (typeof value !== 'string') fail(`op ${name}: text must be a string`)
      if ((value as string).length > OPS_TEXT_MAX_CHARS) {
        fail(
          `op ${name}: text is ${String((value as string).length)} characters; the cap is ${String(
            OPS_TEXT_MAX_CHARS,
          )}`,
        )
      }
      return value as string
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
          const text = takeText(name, op.text, true)
          if (text.length === 0) fail('op insertLines: text must be non-empty')
          const eol = dominantEol(work)
          const shapeNote = MarkdownSession.ensureTerminatedBefore(work, after as number, eol)
          // same trailing-newline preservation as insert_content
          const inserted = toLines(text, eol, shapeNote !== null ? '' : eol)
          work.splice((after as number) + 1, 0, ...inserted)
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
          const matchCase = op.matchCase !== false
          const from = op.from === undefined ? 0 : op.from
          const to = op.to === undefined ? lineCount() - 1 : op.to
          checkRange(name, from, to)
          const needle = matchCase ? find : find.toLowerCase()
          let occurrences = 0
          let changedLines = 0
          for (let i = from as number; i <= (to as number); i++) {
            const hay = work[i]!
            const subject = matchCase ? hay.text : hay.text.toLowerCase()
            if (!subject.includes(needle)) continue
            let replaced: string
            if (matchCase) {
              occurrences += hay.text.split(find).length - 1
              replaced = hay.text.split(find).join(replace)
            } else {
              // rebuild case-insensitively: walk the lowered subject
              let out = ''
              let rest = hay.text
              let restLower = subject
              for (;;) {
                const at = restLower.indexOf(needle)
                if (at === -1) break
                out += rest.slice(0, at) + replace
                rest = rest.slice(at + needle.length)
                restLower = restLower.slice(at + needle.length)
                occurrences += 1
              }
              replaced = out + rest
            }
            work[i] = { ...hay, text: replaced }
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
            `unknown op "${name}" - markdown sessions accept insertLines, replaceLines, ` +
              'deleteLines, findReplace',
          )
      }
    }
    if (!dryRun) {
      this.lines = work
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
   * with the original BOM flag re-applied.
   */
  async save(rawPath?: string, options: { overwrite?: boolean } = {}): Promise<MarkdownSaveResult> {
    const target = resolveConfined(rawPath ?? this.path)
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
      bytes = new TextEncoder().encode(joinLines(this.lines, this.bom))
      if (this.encoding !== 'utf-8') {
        warnings.push(`Original encoding was ${this.encoding}; the edited copy is saved as UTF-8.`)
      }
    }
    await mkdir(dirname(target), { recursive: true })
    const tmp = join(
      dirname(target),
      `.${target.split('/').pop() ?? 'markdown'}.airy-${randomUUID()}`,
    )
    await writeFile(tmp, bytes)
    if (options.overwrite === true || target === this.path || this.savedTargets.has(target)) {
      await rename(tmp, target)
    } else {
      await promoteNewFileExclusively(tmp, target)
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

  /** Markdown sessions hold no external resources. */
  close(): Promise<string[]> {
    return Promise.resolve([])
  }
}

/** clip long read output at the MCP answer budget, with an actionable note */
function clip(text: string, max: number, hint: string): string {
  return text.length <= max
    ? text
    : `${text.slice(0, max)}\n...(output truncated at ${String(max)} characters; ${hint})`
}

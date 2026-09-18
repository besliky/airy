// Headless Markdown editing session (PAR-003): one open .md/.markdown file
// -> line-oriented in-memory model -> byte-preserving save. The line machinery
// (BOM/EOL policy, open caps, line ops, atomic save) lives in the shared
// sessions/line-core.ts (REFR-001, extracted from the PAR-003/PAR-004 twins);
// this file is the markdown half of the contract: the UTF-8-only charset
// policy, the ATX heading structure model, and the read/meta rendering.
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
import { basename } from 'node:path'

import { countWords } from '../docx/session.js'
import {
  decodeStrict,
  dominantEol,
  HEADING_LIST_MAX,
  INSERT_MAX_CHARS,
  LineDocument,
  type Line,
  type LineInsertResult,
  type LineOpResult,
  type LineReadOptions,
  type LineSaveResult,
  type LineSessionHooks,
  renderRead,
  sniffBom,
} from '../sessions/line-core.js'

/** per-heading text width in the read summary (megabyte heading lines stay bounded) */
const HEADING_TEXT_MAX_CHARS = 80

export type MarkdownLine = Line

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

export type MarkdownReadOptions = LineReadOptions

export type MarkdownSaveResult = LineSaveResult

export type InsertResult = LineInsertResult

export type MarkdownOpResult = LineOpResult

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

/**
 * Decode markdown bytes: a BOM picks its charset and survives as a returned
 * flag; BOM-less files must be valid UTF-8. Returns null (not an exception)
 * so open() can phrase one actionable error naming the file.
 */
function decodeMarkdown(
  bytes: Uint8Array,
): { text: string; bom: boolean; encoding: 'utf-8' | 'utf-16le' | 'utf-16be' } | null {
  const sniff = sniffBom(bytes)
  if (sniff.status === 'decoded') {
    return { text: sniff.text, bom: true, encoding: sniff.encoding }
  }
  if (sniff.status === 'none') {
    const text = decodeStrict(bytes, 'utf-8')
    if (text !== null) return { text, bom: false, encoding: 'utf-8' }
  }
  return null
}

/** the markdown divergence points of the shared line-session core */
const markdownHooks: LineSessionHooks = {
  kind: 'markdown',
  decode: decodeMarkdown,
  refusalMessage: (path) =>
    `Cannot open "${path}": the bytes are not valid UTF-8 (markdown is a UTF-8 format; ` +
    'BOM-prefixed UTF-8/UTF-16 is accepted). Convert the file and retry.',
  insertOpText: (value: unknown, fail: (message: string) => never): string => {
    if (value === undefined || value === null) fail('op insertLines: text is required')
    if (typeof value !== 'string') fail('op insertLines: text must be a string')
    if (value.length === 0) fail('op insertLines: text must be non-empty')
    return value
  },
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

  private readonly doc: LineDocument

  private constructor(handle: string, doc: LineDocument) {
    this.handle = handle
    this.path = doc.path
    this.doc = doc
  }

  /** Open a .md/.markdown file inside the workspace root. */
  static async open(rawPath: string, root?: string): Promise<MarkdownSession> {
    const { handle, doc } = await LineDocument.open(rawPath, root, markdownHooks)
    return new MarkdownSession(handle, doc)
  }

  // ---- reading ----

  meta(): MarkdownSessionMeta {
    const text = this.doc.text
    return {
      handle: this.handle,
      kind: 'markdown',
      path: this.path,
      // path.basename is platform-aware: on Windows a `\`-separated path
      // never split on '/', which made the label the whole path (BUG-706)
      fileName: basename(this.path) || this.path,
      format: 'md',
      converted: false,
      editable: true,
      warnings:
        this.doc.encoding === 'utf-8'
          ? []
          : [`Original encoding is UTF-16 (${this.doc.encoding}); saving an edit writes UTF-8.`],
      lineCount: this.doc.lineCount,
      headingCount: scanHeadings(this.doc.lines).length,
      wordCount: countWords(text),
      charCount: text.length,
      eol: dominantEol(this.doc.lines),
      bom: this.doc.bom,
      dirty: this.doc.isDirty,
    }
  }

  private headings(): MarkdownHeading[] {
    return scanHeadings(this.doc.lines)
  }

  /**
   * Agent-facing read: file header (stats + EOL/BOM), the heading list with
   * line positions, then the full text (the shared read tail applies the 30k
   * budget and the lines/range selection).
   */
  readDocument(options: MarkdownReadOptions = {}): string {
    const meta = this.meta()
    const header =
      `"${meta.fileName}" - ${String(meta.lineCount)} lines, ${String(meta.wordCount)} words, ` +
      `${String(meta.charCount)} characters. Line endings: ${
        meta.eol === '\r\n' ? 'CRLF' : meta.eol === '\r' ? 'CR' : 'LF'
      }.${meta.bom ? ' Starts with a BOM.' : ''}`
    const headings = this.headings()
    const headingLines = headings.map(
      (h) => `${h.ordinal}|${h.line}|${'#'.repeat(h.level)}|${clipText(h.text)}`,
    )
    const headingsBlock = [
      `Headings (ordinal|line|level|text):${headingLines.length === 0 ? ' none' : ''}${
        headings.length > HEADING_LIST_MAX
          ? ` (first ${String(HEADING_LIST_MAX)} of ${String(headings.length)} - use range reads for the rest)`
          : ''
      }`,
      ...headingLines.slice(0, HEADING_LIST_MAX),
    ].join('\n')
    return renderRead(header, headingsBlock, this.doc.lines, options)
  }

  // ---- editing ----

  /**
   * Resolve an afterHeading ordinal (markdown's extra insert anchor) to its
   * line, with the re-read hint on a stale ordinal.
   */
  private headingInsertLine(afterHeading: number): { ordinal: number; line: number } {
    const headings = this.headings()
    const ordinal = Math.trunc(afterHeading)
    if (ordinal < 1 || ordinal > headings.length) {
      throw new Error(
        `afterHeading ${String(ordinal)} is out of range (the document has ${String(
          headings.length,
        )} heading(s)) - re-read the document for the current ordinals`,
      )
    }
    return { ordinal, line: headings[ordinal - 1]!.line }
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
    // the heading anchor only resolves when marker does not take precedence,
    // matching the original marker > afterHeading > at order
    const afterHeading =
      position.marker === undefined && position.afterHeading !== undefined
        ? this.headingInsertLine(position.afterHeading)
        : undefined
    return this.doc.insert(text, position, afterHeading !== undefined ? { afterHeading } : {})
  }

  /** Apply a batch of line ops atomically (validated/applied on a copy). */
  applyOps(
    ops: Array<Record<string, unknown>>,
    dryRun = false,
  ): { results: MarkdownOpResult[]; summary: string; dryRun: boolean } {
    return this.doc.applyOps(ops, dryRun)
  }

  // ---- saving ----

  /**
   * Save atomically via the shared core (tmp + rename, drift refusal, target
   * ownership): with no edits the original bytes round-trip verbatim; an
   * edited save writes UTF-8 with the original BOM flag re-applied.
   */
  async save(rawPath?: string, options: { overwrite?: boolean } = {}): Promise<MarkdownSaveResult> {
    return this.doc.save(rawPath, options)
  }

  /** Markdown sessions hold no external resources. */
  close(): Promise<string[]> {
    return Promise.resolve([])
  }
}

/** cap one heading's text in the read summary (heading lines can be megabytes) */
function clipText(text: string, max = HEADING_TEXT_MAX_CHARS): string {
  return text.length <= max ? text : `${text.slice(0, max - 1)}...`
}

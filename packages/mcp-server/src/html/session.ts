// Headless HTML editing session (PAR-004): one open .html/.htm file ->
// line-oriented in-memory model -> byte-preserving save, the same contract the
// markdown session (PAR-003) established for plain-text formats. The file is
// split into lines that each keep their OWN terminator, edits splice that
// array, and untouched lines (terminators included) round-trip
// byte-identically on save; a zero-edit save writes the original bytes back
// verbatim, so a file the agent only read never changes on disk.
//
// Structure model for read_document: a parse5 pass (the parser the suite
// already ships - apps/html builds its parse map on it) reports h1-h6
// headings and a[href] links with their line positions, plus the <title>;
// inserted fragments are spliced VERBATIM (no reparse/rewrite), so what the
// agent sends is exactly what lands on disk.
//
// Encoding policy (mirrors PAR-003 / CHANGELOG 0.9.3 BOM semantics): UTF-8
// with or without a BOM and BOM-prefixed UTF-16 are accepted natively; bytes
// that are not valid UTF-8 open only when the document itself declares a
// usable legacy charset (<meta charset>, the signal browsers trust - the
// file-parse decodeHtmlText approach minus its silent generic fallback,
// which would bless mojibake with a save). A leading BOM lives in a flag,
// stays out of the editable text, and is re-applied on every save; edited
// saves always write UTF-8, and a legacy charset declaration is rewritten to
// utf-8 so the re-encoded file renders correctly in browsers.
import { randomUUID } from 'node:crypto'
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'

import { parse as parseHtml5, type DefaultTreeAdapterTypes as T } from 'parse5'

import {
  assertSaveTargetFree,
  countWords,
  FencingError,
  promoteNewFileExclusively,
} from '../docx/session.js'
import { resolveConfined } from '../docx/paths.js'

// ---- limits (mirror the docx/markdown sessions, scaled to the MCP 30k answer budget) ----

const MAX_OPEN_BYTES = 8 * 1024 * 1024
const READ_MAX_CHARS = 30_000
const INSERT_MAX_CHARS = 200_000
const OPS_TEXT_MAX_CHARS = 200_000
/** files above this skip the parse5 structure scan (read shows the text only) */
const STRUCTURE_SCAN_MAX_CHARS = 1_000_000
/** how many links the read summary lists before eliding */
const LINK_LIST_MAX = 200
/**
 * Largest range span a read materializes: the request schema does not bound
 * `end`, so the session must reject a huge span BEFORE building the index
 * array (a range like 0..2^53 would otherwise hang/OOM the server).
 */
const RANGE_MAX_SPAN = 10_000

/** U+FEFF as pure-ASCII source (a literal BOM char in source trips tooling) */
const BOM_CHAR = String.fromCharCode(0xfeff)
const REPLACEMENT_CHAR = String.fromCharCode(0xfffd)
const NUL_CHAR = String.fromCharCode(0)

export interface HtmlLine {
  text: string
  /** '\n' | '\r\n' | '\r' for terminated lines, '' for the final unterminated one */
  eol: string
}

export interface HtmlHeading {
  /** 1-based ordinal (informational addressing aid for insert markers) */
  ordinal: number
  /** 0-based line index of the opening tag */
  line: number
  level: number
  text: string
}

export interface HtmlLink {
  ordinal: number
  /** 0-based line index of the opening tag */
  line: number
  href: string
  text: string
}

export interface HtmlStructure {
  title: string | null
  headings: HtmlHeading[]
  links: HtmlLink[]
}

export interface HtmlSessionMeta {
  handle: string
  kind: 'html'
  path: string
  fileName: string
  format: 'html'
  converted: false
  editable: true
  warnings: string[]
  lineCount: number
  headingCount: number
  linkCount: number
  title: string | null
  wordCount: number
  charCount: number
  /** dominant line ending: what inserted/edited lines get */
  eol: '\n' | '\r\n' | '\r'
  /** the file started with a byte order mark */
  bom: boolean
  dirty: boolean
}

export interface HtmlReadOptions {
  /** line indexes to return in full; must exist (read_document blocks) */
  lines?: number[]
  /** inclusive line range to return in full */
  range?: { start: number; end: number }
}

export interface HtmlSaveResult {
  path: string
  bytes: number
  unchanged: boolean
  warnings: string[]
}

export interface HtmlInsertResult {
  /** lines inserted */
  inserted: number
  /** line index the fragment landed after (-1 = document start) */
  at: number
  lineCount: number
  dirty: boolean
  /** human-readable position summary */
  detail: string
}

export interface HtmlOpResult {
  op: string
  matched: number
  changed: number
  detail?: string
}

interface FileStamp {
  mtimeMs: number
  size: number
}

// ---- text <-> line model (same contract as the markdown session) ----

/** Split text into lines, each remembering its own terminator. */
export function splitLines(text: string): HtmlLine[] {
  const lines: HtmlLine[] = []
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
function dominantEol(lines: readonly HtmlLine[]): '\n' | '\r\n' | '\r' {
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
function joinLines(lines: readonly HtmlLine[], bom: boolean): string {
  return (bom ? BOM_CHAR : '') + lines.map((l) => l.text + l.eol).join('')
}

/**
 * Agent text -> session lines. `lastEol` is what the final line of the block
 * carries: the dominant EOL for a spliced-in block, or '' to preserve a file
 * end that had no trailing newline.
 */
function toLines(text: string, eol: string, lastEol: string): HtmlLine[] {
  const parts = text.split(/\r\n|\r|\n/)
  return parts.map((t, i) => ({ text: t, eol: i === parts.length - 1 ? lastEol : eol }))
}

// ---- structure scan (parse5, same parser apps/html builds on) ----

function isElement(node: T.Node): node is T.Element {
  return 'tagName' in node && typeof (node as T.Element).tagName === 'string'
}

/** concatenated text of an element's descendants (script/style excluded) */
function textContent(node: T.Node): string {
  if ('value' in node && typeof (node as T.TextNode).value === 'string') {
    return (node as T.TextNode).value
  }
  if ('childNodes' in node) {
    let out = ''
    for (const child of node.childNodes) {
      if (isElement(child) && (child.tagName === 'script' || child.tagName === 'style')) continue
      out += textContent(child)
    }
    return out
  }
  return ''
}

function collapse(text: string): string {
  return text.replace(/\s+/g, ' ').trim()
}

/**
 * Headings, links and the title with 0-based line positions, in document
 * order. Elements are deduped by start-tag offset (the adoption agency can
 * re-create formatting elements sharing one start tag). Oversize documents
 * skip the scan: pass null text to signal "structure unavailable".
 */
export function scanStructure(text: string | null): HtmlStructure {
  if (text === null || text.length > STRUCTURE_SCAN_MAX_CHARS) {
    return { title: null, headings: [], links: [] }
  }
  const doc = parseHtml5(text, { sourceCodeLocationInfo: true })
  const headings: HtmlHeading[] = []
  const links: HtmlLink[] = []
  let title: string | null = null
  const seen = new Set<number>()
  const walk = (node: T.Node) => {
    if (isElement(node)) {
      const loc = node.sourceCodeLocation
      const startOffset = loc?.startTag?.startOffset ?? loc?.startOffset
      const headingLevel = /^h([1-6])$/.exec(node.tagName)
      const isLink = node.tagName === 'a'
      const isTitle = node.tagName === 'title' && title === null
      if (startOffset !== undefined && !seen.has(startOffset)) {
        seen.add(startOffset)
        const line = Math.max(0, (loc!.startLine ?? 1) - 1)
        if (headingLevel) {
          headings.push({
            ordinal: headings.length + 1,
            line,
            level: Number(headingLevel[1]),
            text: clipText(collapse(textContent(node))),
          })
        } else if (isLink) {
          const href = node.attrs.find((a) => a.name === 'href')?.value ?? ''
          links.push({
            ordinal: links.length + 1,
            line,
            href: clipText(href),
            text: clipText(collapse(textContent(node))),
          })
        } else if (isTitle) {
          title = collapse(textContent(node))
        }
      }
    }
    if ('childNodes' in node) for (const child of node.childNodes) walk(child)
    if (isElement(node) && node.tagName === 'template') {
      const content = (node as T.Template).content
      if (content) for (const child of content.childNodes) walk(child)
    }
  }
  walk(doc)
  return { title, headings, links }
}

function clipText(text: string, max = 80): string {
  return text.length <= max ? text : `${text.slice(0, max - 1)}…`
}

// ---- decode (BOM-aware strict UTF-8/UTF-16, declared charsets, no generic fallback) ----

function decodeStrict(bytes: Uint8Array, charset: string): string | null {
  try {
    return new TextDecoder(charset, { fatal: true, ignoreBOM: true }).decode(bytes)
  } catch {
    return null
  }
}

const META_CHARSET_RE = /<meta[^>]+charset\s*=\s*["']?\s*([a-z0-9_.:-]+)/i

/**
 * Decode HTML bytes: BOM picks its charset and survives as a flag; BOM-less
 * files must be valid UTF-8, or declare a legacy charset the document itself
 * carries (browsers trust <meta charset> too). Returns null so open() can
 * phrase one actionable error naming the file.
 */
function decodeHtml(bytes: Uint8Array): { text: string; bom: boolean; encoding: string } | null {
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
  const utf8 = decodeStrict(bytes, 'utf-8')
  if (utf8 !== null) return { text: utf8, bom: false, encoding: 'utf-8' }
  // windows-1252 maps every byte, so the ASCII head shows verbatim for the sniff
  const head = decodeStrict(bytes.subarray(0, 4096), 'windows-1252') ?? ''
  const declared = META_CHARSET_RE.exec(head)?.[1]
  if (declared !== undefined && !/^utf-?8$/i.test(declared)) {
    try {
      const text = new TextDecoder(declared, { fatal: true }).decode(bytes)
      if (!text.includes(REPLACEMENT_CHAR)) return { text, bom: false, encoding: declared }
    } catch {
      // declared charset cannot decode the bytes - fall through to refusal
    }
  }
  return null
}

// ---- insert position params ----

export interface HtmlInsertPosition {
  /** insert after this line index (-1 = document start) */
  at?: number
  /** insert after the first line containing this exact substring */
  marker?: string
}

// ---- session ----

export class HtmlSession {
  readonly handle: string
  readonly path: string

  private readonly originalBytes: Uint8Array
  private lines: HtmlLine[]
  private readonly bom: boolean
  private readonly encoding: string
  private baseline: FileStamp | null
  private readonly savedTargets = new Set<string>()
  private dirty = false

  private constructor(
    handle: string,
    path: string,
    originalBytes: Uint8Array,
    lines: HtmlLine[],
    bom: boolean,
    encoding: string,
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

  /** Open a .html/.htm file inside the workspace root. */
  static async open(rawPath: string, root?: string): Promise<HtmlSession> {
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
        `"${path}" is ${String(bytes.byteLength)} bytes; html sessions cap at ` +
          `${String(MAX_OPEN_BYTES)} bytes (8 MiB).`,
      )
    }
    const decoded = decodeHtml(bytes)
    if (decoded === null) {
      throw new Error(
        `Cannot open "${path}": the bytes are not valid UTF-8 and the document declares no ` +
          'usable charset (HTML sessions accept UTF-8, BOM-prefixed UTF-8/UTF-16, and declared ' +
          'legacy charsets). Convert the file and retry.',
      )
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
    return new HtmlSession(
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

  meta(): HtmlSessionMeta {
    const text = joinLines(this.lines, false)
    const structure = this.structure()
    return {
      handle: this.handle,
      kind: 'html',
      path: this.path,
      fileName: this.path.split('/').pop() ?? this.path,
      format: 'html',
      converted: false,
      editable: true,
      warnings: decodedIsUtf8(this.encoding)
        ? []
        : [
            `Original encoding is ${this.encoding}; saving an edit writes UTF-8 ` +
              '(zero-edit saves keep the original bytes).',
          ],
      lineCount: this.lines.length,
      headingCount: structure.headings.length,
      linkCount: structure.links.length,
      title: structure.title,
      wordCount: countWords(text),
      charCount: text.length,
      eol: dominantEol(this.lines),
      bom: this.bom,
      dirty: this.dirty,
    }
  }

  private structure(): HtmlStructure {
    const text = joinLines(this.lines, false)
    return scanStructure(text.length > STRUCTURE_SCAN_MAX_CHARS ? null : text)
  }

  /**
   * Agent-facing read: file header (title, stats, EOL/BOM), the structural
   * summary (headings and links with line positions), then the full text -
   * truncated at the 30k budget with a hint to request line ranges instead.
   * With lines/range selected, the full text section is replaced by exactly
   * those lines.
   */
  readDocument(options: HtmlReadOptions = {}): string {
    const meta = this.meta()
    const header =
      `"${meta.fileName}" - ${String(meta.lineCount)} lines, ${String(meta.wordCount)} words, ` +
      `${String(meta.charCount)} characters.` +
      `${meta.title ? ` Title: ${meta.title}.` : ''}` +
      ` Line endings: ${meta.eol === '\r\n' ? 'CRLF' : meta.eol === '\r' ? 'CR' : 'LF'}.` +
      `${meta.bom ? ' Starts with a BOM.' : ''}`
    const structure = this.structure()
    const oversize = joinLines(this.lines, false).length > STRUCTURE_SCAN_MAX_CHARS
    const headingLines = structure.headings.map(
      (h) => `${h.ordinal}|${h.line}|h${h.level}|${h.text}`,
    )
    const linkLines = structure.links
      .slice(0, LINK_LIST_MAX)
      .map((l) => `${l.ordinal}|${l.line}|${l.text || '(no text)'} -> ${l.href}`)
    const structureBlock = oversize
      ? 'Structure (ordinal|line|tag|text): skipped - the document exceeds the 1M-character scan budget; address lines directly.'
      : [
          `Headings (ordinal|line|level|text):${headingLines.length === 0 ? ' none' : ''}`,
          ...headingLines,
          `Links (ordinal|line|text -> href):${
            structure.links.length === 0 ? ' none' : ''
          }${structure.links.length > LINK_LIST_MAX ? ` (first ${String(LINK_LIST_MAX)} of ${String(structure.links.length)})` : ''}`,
          ...linkLines,
        ].join('\n')

    const selected = this.selectedIndexes(options)
    if (selected === null) {
      const fullText = this.lines.map((l) => l.text).join('\n')
      const clipped = clip(fullText, READ_MAX_CHARS, 'request a line range to read the rest')
      return [header, structureBlock, '', 'Full text (EOLs normalized to LF):', clipped].join('\n')
    }
    const body = selected.map((i) => this.lines[i]!.text).join('\n')
    const clipped = clip(body, READ_MAX_CHARS, 'request a narrower selection')
    return [
      header,
      structureBlock,
      '',
      `Selected ${String(selected.length)} line(s) (EOLs normalized to LF):`,
      clipped,
    ].join('\n')
  }

  /** Resolve blocks/range-style selections to validated, sorted line indexes. */
  private selectedIndexes(options: HtmlReadOptions): number[] | null {
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
    lines: HtmlLine[],
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
   * Insert an HTML fragment verbatim: after the first line containing
   * `marker` (e.g. '</body>' to append rendered content), after line `at`
   * (-1 = document start), or at the end by default. Precedence: marker > at.
   */
  insertContent(html: string, position: HtmlInsertPosition = {}): HtmlInsertResult {
    if (typeof html !== 'string' || html.length === 0) {
      throw new Error('html is required for an html insert (the fragment is spliced verbatim)')
    }
    if (html.length > INSERT_MAX_CHARS) {
      throw new Error(
        `insert html is ${String(html.length)} characters; the cap is ${String(INSERT_MAX_CHARS)}.`,
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
    const shapeNote = HtmlSession.ensureTerminatedBefore(this.lines, after, eol)
    // appending at the end preserves the file's trailing-newline shape: a
    // file that ended without a newline keeps ending without one (shapeNote
    // firing means exactly that case)
    const newLines = toLines(html, eol, shapeNote !== null ? '' : eol)
    this.lines.splice(after + 1, 0, ...newLines)
    this.dirty = true
    const detail = `inserted ${String(newLines.length)} line(s) ${where}${
      shapeNote ? ` (${shapeNote})` : ''
    } verbatim`
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
  ): { results: HtmlOpResult[]; summary: string; dryRun: boolean } {
    const work = this.lines.map((l) => ({ ...l }))
    const results: HtmlOpResult[] = []
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
          const text =
            typeof op.text === 'string' && op.text.length > 0
              ? op.text
              : fail('op insertLines: text is required (non-empty HTML/markup source)')
          if (text.length > OPS_TEXT_MAX_CHARS) {
            fail(
              `op insertLines: text is ${String(text.length)} characters; the cap is ${String(
                OPS_TEXT_MAX_CHARS,
              )}`,
            )
          }
          const eol = dominantEol(work)
          const shapeNote = HtmlSession.ensureTerminatedBefore(work, after as number, eol)
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
            `unknown op "${name}" - html sessions accept insertLines, replaceLines, ` +
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
   * Point the document's charset declaration at UTF-8 before an edited save
   * writes UTF-8 bytes: browsers trust <meta charset>, so a stale legacy label
   * would decode the re-encoded file as mojibake (the GUI editor follows the
   * same decode-declared / save-UTF-8 policy). Returns the replaced label, or
   * null when nothing needed rewriting. The charset token itself cannot span
   * lines, so the joined-text offset maps onto exactly one line.
   */
  private pinCharsetDeclaration(): string | null {
    const text = this.lines.map((l) => l.text).join('\n')
    const match = META_CHARSET_RE.exec(text)
    if (match === null || /^utf-?8$/i.test(match[1]!)) return null
    const tokenStart = match.index + match[0].length - match[1]!.length
    let offset = 0
    for (const line of this.lines) {
      if (tokenStart >= offset && tokenStart < offset + line.text.length) {
        const at = tokenStart - offset
        line.text = `${line.text.slice(0, at)}utf-8${line.text.slice(at + match[1]!.length)}`
        return match[1]!
      }
      offset += line.text.length + 1
    }
    return null
  }

  /**
   * Save atomically (tmp + rename) with the docx session's fences: saving over
   * the opened file refuses when it changed on disk since open; a target that
   * exists is refused unless the session owns it or overwrite is true. With no
   * edits the original bytes round-trip verbatim (an untouched file never
   * changes on disk, whatever its encoding was); an edited save writes UTF-8
   * with the original BOM flag re-applied, and a legacy charset declaration is
   * rewritten to utf-8 so the saved file decodes correctly in browsers.
   */
  async save(rawPath?: string, options: { overwrite?: boolean } = {}): Promise<HtmlSaveResult> {
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
      if (!decodedIsUtf8(this.encoding)) {
        const replaced = this.pinCharsetDeclaration()
        if (replaced !== null) {
          warnings.push(
            `Charset declaration rewritten from "${replaced}" to "utf-8" ` +
              '(the edited copy is UTF-8; a stale legacy claim would render as mojibake).',
          )
        }
      }
      bytes = new TextEncoder().encode(joinLines(this.lines, this.bom))
      if (!decodedIsUtf8(this.encoding)) {
        warnings.push(`Original encoding was ${this.encoding}; the edited copy is saved as UTF-8.`)
      }
    }
    await mkdir(dirname(target), { recursive: true })
    const tmp = join(dirname(target), `.${target.split('/').pop() ?? 'html'}.airy-${randomUUID()}`)
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

  /** HTML sessions hold no external resources. */
  close(): Promise<string[]> {
    return Promise.resolve([])
  }
}

/** true for the encodings whose edited saves stay byte-comparable to UTF-8 */
function decodedIsUtf8(encoding: string): boolean {
  return encoding === 'utf-8'
}

/** clip long read output at the MCP answer budget, with an actionable note */
function clip(text: string, max: number, hint: string): string {
  return text.length <= max
    ? text
    : `${text.slice(0, max)}\n...(output truncated at ${String(max)} characters; ${hint})`
}

// Headless HTML editing session (PAR-004): one open .html/.htm file ->
// line-oriented in-memory model -> byte-preserving save, the same contract the
// markdown session (PAR-003) established for plain-text formats. The line
// machinery (BOM/EOL policy, open caps, line ops, atomic save) lives in the
// shared sessions/line-core.ts (REFR-001, extracted from the PAR-003/PAR-004
// twins); this file is the html half of the contract: the parse5 structure
// model, the declared-charset decode policy, and the read/meta rendering.
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
import { basename } from 'node:path'

import { parse as parseHtml5, type DefaultTreeAdapterTypes as T } from 'parse5'

import { countWords } from '../docx/session.js'
import {
  decodeStrict,
  dominantEol,
  HEADING_LIST_MAX,
  INSERT_MAX_CHARS,
  LineDocument,
  type DecodedText,
  type Line,
  type LineInsertResult,
  type LineOpResult,
  type LineReadOptions,
  type LineSaveResult,
  type LineSessionHooks,
  renderRead,
  sniffBom,
} from '../sessions/line-core.js'

/** files above this skip the parse5 structure scan (read shows the text only) */
const STRUCTURE_SCAN_MAX_CHARS = 1_000_000
/** how many links the read summary lists before eliding */
const LINK_LIST_MAX = 200

/**
 * Closing tags whose line a marker insert must land BEFORE: a marker like
 * "</body>" names the element boundary, and splicing after the line would
 * push the fragment between </body> and </html> — outside the body element
 * (DOC-1506). Only a line that is JUST the closing tag flips (a line that
 * carries other markup keeps the documented after-the-line behavior).
 */
const STRUCTURAL_CLOSING_TAGS = new Set(['</body>', '</html>', '</head>'])

function isStructuralClosingLine(lineText: string): boolean {
  return STRUCTURAL_CLOSING_TAGS.has(lineText.trim().toLowerCase())
}

const REPLACEMENT_CHAR = String.fromCharCode(0xfffd)

export type HtmlLine = Line

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

export type HtmlReadOptions = LineReadOptions

export type HtmlSaveResult = LineSaveResult

export type HtmlInsertResult = LineInsertResult

export type HtmlOpResult = LineOpResult

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

const META_CHARSET_RE = /<meta[^>]+charset\s*=\s*["']?\s*([a-z0-9_.:-]+)/i

/**
 * Decode HTML bytes: BOM picks its charset and survives as a flag; BOM-less
 * files must be valid UTF-8, or declare a legacy charset the document itself
 * carries (browsers trust <meta charset> too). Returns null so open() can
 * phrase one actionable error naming the file.
 */
function decodeHtml(bytes: Uint8Array): DecodedText | null {
  const sniff = sniffBom(bytes)
  if (sniff.status === 'decoded') {
    return { text: sniff.text, bom: true, encoding: sniff.encoding }
  }
  if (sniff.status === 'none') {
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
  }
  return null
}

/** true for the encodings whose edited saves stay byte-comparable to UTF-8 */
function decodedIsUtf8(encoding: string): boolean {
  return encoding === 'utf-8'
}

/**
 * Point the document's charset declaration at UTF-8 before an edited save
 * writes UTF-8 bytes: browsers trust <meta charset>, so a stale legacy label
 * would decode the re-encoded file as mojibake (the GUI editor follows the
 * same decode-declared / save-UTF-8 policy). Returns the replaced label, or
 * null when nothing needed rewriting. The charset token itself cannot span
 * lines, so the joined-text offset maps onto exactly one line.
 */
function pinCharsetDeclaration(lines: Line[]): string | null {
  const text = lines.map((l) => l.text).join('\n')
  const match = META_CHARSET_RE.exec(text)
  if (match === null || /^utf-?8$/i.test(match[1]!)) return null
  const tokenStart = match.index + match[0].length - match[1]!.length
  let offset = 0
  for (const line of lines) {
    if (tokenStart >= offset && tokenStart < offset + line.text.length) {
      const at = tokenStart - offset
      line.text = `${line.text.slice(0, at)}utf-8${line.text.slice(at + match[1]!.length)}`
      return match[1]!
    }
    offset += line.text.length + 1
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

/** the html divergence points of the shared line-session core */
const htmlHooks: LineSessionHooks = {
  kind: 'html',
  decode: decodeHtml,
  refusalMessage: (path) =>
    `Cannot open "${path}": the bytes are not valid UTF-8 and the document declares no ` +
    'usable charset (HTML sessions accept UTF-8, BOM-prefixed UTF-8/UTF-16, and declared ' +
    'legacy charsets). Convert the file and retry.',
  insertOpText: (value: unknown, fail: (message: string) => never): string => {
    if (typeof value !== 'string' || value.length === 0) {
      fail('op insertLines: text is required (non-empty HTML/markup source)')
    }
    return value
  },
  beforeEncode: (lines, warnings, encoding) => {
    if (decodedIsUtf8(encoding)) return
    const replaced = pinCharsetDeclaration(lines)
    if (replaced !== null) {
      warnings.push(
        `Charset declaration rewritten from "${replaced}" to "utf-8" ` +
          '(the edited copy is UTF-8; a stale legacy claim would render as mojibake).',
      )
    }
  },
  insertBeforeMarkerLine: isStructuralClosingLine,
}

// ---- session ----

export class HtmlSession {
  readonly handle: string
  readonly path: string

  private readonly doc: LineDocument

  private constructor(handle: string, doc: LineDocument) {
    this.handle = handle
    this.path = doc.path
    this.doc = doc
  }

  /** Open a .html/.htm file inside the workspace root. */
  static async open(rawPath: string, root?: string): Promise<HtmlSession> {
    const { handle, doc } = await LineDocument.open(rawPath, root, htmlHooks)
    return new HtmlSession(handle, doc)
  }

  // ---- reading ----

  meta(): HtmlSessionMeta {
    const text = this.doc.text
    const structure = this.structure()
    return {
      handle: this.handle,
      kind: 'html',
      path: this.path,
      // path.basename is platform-aware: on Windows a `\`-separated path
      // never split on '/', which made the label the whole path (BUG-706)
      fileName: basename(this.path) || this.path,
      format: 'html',
      converted: false,
      editable: true,
      warnings: decodedIsUtf8(this.doc.encoding)
        ? []
        : [
            `Original encoding is ${this.doc.encoding}; saving an edit writes UTF-8 ` +
              '(zero-edit saves keep the original bytes).',
          ],
      lineCount: this.doc.lineCount,
      headingCount: structure.headings.length,
      linkCount: structure.links.length,
      title: structure.title,
      wordCount: countWords(text),
      charCount: text.length,
      eol: dominantEol(this.doc.lines),
      bom: this.doc.bom,
      dirty: this.doc.isDirty,
    }
  }

  private structure(): HtmlStructure {
    const text = this.doc.text
    return scanStructure(text.length > STRUCTURE_SCAN_MAX_CHARS ? null : text)
  }

  /**
   * Agent-facing read: file header (title, stats, EOL/BOM), the structural
   * summary (headings and links with line positions - both lists capped),
   * then the full text (the shared read tail applies the 30k budget and the
   * lines/range selection).
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
    const oversize = this.doc.text.length > STRUCTURE_SCAN_MAX_CHARS
    const headingLines = structure.headings
      .slice(0, HEADING_LIST_MAX)
      .map((h) => `${h.ordinal}|${h.line}|h${h.level}|${h.text}`)
    const linkLines = structure.links
      .slice(0, LINK_LIST_MAX)
      .map((l) => `${l.ordinal}|${l.line}|${l.text || '(no text)'} -> ${l.href}`)
    const structureBlock = oversize
      ? 'Structure (ordinal|line|tag|text): skipped - the document exceeds the 1M-character scan budget; address lines directly.'
      : [
          `Headings (ordinal|line|level|text):${structure.headings.length === 0 ? ' none' : ''}${
            structure.headings.length > HEADING_LIST_MAX
              ? ` (first ${String(HEADING_LIST_MAX)} of ${String(
                  structure.headings.length,
                )} - use range reads for the rest)`
              : ''
          }`,
          ...headingLines,
          `Links (ordinal|line|text -> href):${
            structure.links.length === 0 ? ' none' : ''
          }${structure.links.length > LINK_LIST_MAX ? ` (first ${String(LINK_LIST_MAX)} of ${String(structure.links.length)})` : ''}`,
          ...linkLines,
        ].join('\n')
    return renderRead(header, structureBlock, this.doc.lines, options)
  }

  // ---- editing ----

  /**
   * Insert an HTML fragment verbatim: after the first line containing
   * `marker`, after line `at` (-1 = document start), or at the end by
   * default. Precedence: marker > at. A line that is just a closing
   * </body>/</html>/</head> tag inserts BEFORE it, so a marker like
   * "</body>" keeps the fragment inside the element (DOC-1506).
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
    return this.doc.insert(html, position, { detailSuffix: ' verbatim' })
  }

  /** Apply a batch of line ops atomically (validated/applied on a copy). */
  applyOps(
    ops: Array<Record<string, unknown>>,
    dryRun = false,
  ): { results: HtmlOpResult[]; summary: string; dryRun: boolean } {
    return this.doc.applyOps(ops, dryRun)
  }

  // ---- saving ----

  /**
   * Save atomically via the shared core (tmp + rename, drift refusal, target
   * ownership): with no edits the original bytes round-trip verbatim; an
   * edited save writes UTF-8 with the original BOM flag re-applied, and the
   * beforeEncode hook rewrites a legacy charset declaration to utf-8 so the
   * saved file decodes correctly in browsers.
   */
  async save(rawPath?: string, options: { overwrite?: boolean } = {}): Promise<HtmlSaveResult> {
    return this.doc.save(rawPath, options)
  }

  /** HTML sessions hold no external resources. */
  close(): Promise<string[]> {
    return Promise.resolve([])
  }
}

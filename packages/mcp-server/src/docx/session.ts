// Headless docx editing session. One open document -> in-memory engine model
// -> byte-preserving save. This is the headless twin of the renderer's
// PM->saveDocx pipeline, applied straight to the docx-engine Block model:
//
//   open:    parseDocx(bytes) -> ParsedDocFull (blocks with docxIndex anchors)
//   edit:    ops/insert mutate SessionEntry copies ('edited' GeneratedBlocks);
//            'original' entries are never touched, which is what keeps them
//            byte-identical on save
//   save:    saveDocx(parsed, SaveBlock[]) -> tmp file + rename (atomic),
//            with mtime/size fencing against external writers
//
// The renderer's ProseMirror pipeline (2830-line convert.ts + tiptap schema +
// DOMParser-based HTML branch) is not portable without a view, so the ops
// vocabulary here is reimplemented against the engine model directly.
import { randomUUID } from 'node:crypto'
import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'

import {
  parseDocx,
  saveDocx,
  type Block,
  type GeneratedBlock,
  type ParsedDocFull,
  type SaveBlock,
} from '@genoffice/docx-engine'

import { blocksToHtml, blockPreviewText, blockTypeName, parseRestrictedHtml } from './html.js'
import {
  allocateNumId,
  emptyNumbering,
  entryBlock,
  executeOps,
  toSaveBlocks,
  type Op,
  type OpResult,
  type SessionEntry,
} from './ops.js'
import { resolveConfined, workspaceRoot } from './paths.js'

// ---- limits (mirror the embedded agent, scaled to the MCP 30k answer budget) ----

const CONTEXT_MAX_CHARS = 30_000
const PREVIEW_MAX_CHARS = 60
const PREVIEW_TIGHT_CHARS = 20
const READ_MAX_CHARS = 30_000

// Word-parity word count (CJK chars one by one + non-Asian words)
const ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟、-〿぀-ヿ㄀-ㄯ㄰-㆏㇀-ㇿ㐀-䶿一-鿿가-힯豈-﫿！-｠￠-￦]|[\uD840-\uD87F][\uDC00-\uDFFF]/g
const NON_ASIAN_WORD_RE = /[A-Za-z0-9À-ɏ]+(?:['-][A-Za-z0-9À-ɏ]+)*/g

export function countWords(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length + (text.match(NON_ASIAN_WORD_RE) ?? []).length
}

function clip(text: string, max: number): string {
  return text.length <= max ? text : `${text.slice(0, Math.max(1, max - 1))}…`
}

// ---- session ----

export interface SessionMeta {
  handle: string
  path: string
  fileName: string
  blockCount: number
  wordCount: number
  charCount: number
  charCountNoSpaces: number
  dirty: boolean
}

export interface ReadOptions {
  /** block indexes to return in full (restricted HTML); indexes must exist */
  blocks?: number[]
  /** inclusive block range to return in full */
  range?: { start: number; end: number }
}

export interface SaveResult {
  path: string
  bytes: number
  unchanged: boolean
}

export class FencingError extends Error {
  constructor(path: string) {
    super(
      `Refusing to save: "${path}" changed on disk since it was opened ` +
        `(another program may be editing it). Reopen the document and re-apply your edits.`,
    )
    this.name = 'FencingError'
  }
}

interface FileStamp {
  mtimeMs: number
  size: number
}

export class DocxSession {
  readonly handle: string
  /** absolute path the document was opened from */
  readonly path: string

  private readonly root: string
  private readonly parsed: ParsedDocFull
  private entries: SessionEntry[]
  private pending = emptyNumbering()
  private baseline: FileStamp | null
  private savedPath: string | null = null

  private constructor(
    handle: string,
    path: string,
    root: string,
    parsed: ParsedDocFull,
    stamp: FileStamp | null,
  ) {
    this.handle = handle
    this.path = path
    this.root = root
    this.parsed = parsed
    this.entries = DocxSession.initialEntries(parsed)
    this.baseline = stamp
  }

  private static initialEntries(parsed: ParsedDocFull): SessionEntry[] {
    return parsed.blocks
      .filter((block) => !block.hidden)
      .map((block) => ({ kind: 'original' as const, block }))
  }

  /** Open a .docx inside the workspace root and parse it into the session model. */
  static async open(rawPath: string, root?: string): Promise<DocxSession> {
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
    let parsed: ParsedDocFull
    try {
      parsed = await parseDocx(bytes)
    } catch (e) {
      throw new Error(
        `Cannot parse "${path}" as a .docx document: ${e instanceof Error ? e.message : String(e)}`,
        { cause: e },
      )
    }
    return new DocxSession(randomUUID(), path, root ?? workspaceRoot(), parsed, stamp)
  }

  // ---- reading ----

  private view(): Array<Block | GeneratedBlock> {
    return this.entries.map(entryBlock)
  }

  meta(): SessionMeta {
    const fullText = this.entries
      .map((entry) => {
        const b = entryBlock(entry)
        if (b.type === 'paragraph' || b.type === 'heading' || b.type === 'listItem')
          return (b.runs ?? [])
            .filter((r) => !r.del)
            .map((r) => r.text)
            .join('')
        return ''
      })
      .join('')
    return {
      handle: this.handle,
      path: this.path,
      fileName: this.path.split('/').pop() ?? this.path,
      blockCount: this.entries.length,
      wordCount: countWords(fullText),
      charCount: fullText.length,
      charCountNoSpaces: fullText.replace(/\s/g, '').length,
      dirty: this.isDirty(),
    }
  }

  private isDirty(): boolean {
    if (this.savedPath !== null) return true
    return this.entries.some((entry) => entry.kind !== 'original')
  }

  /**
   * Document overview for the agent: `index|type|preview` lines plus stats.
   * With blocks/range selected, the full restricted-HTML content of those
   * blocks is appended instead (same shape the embedded agent's read_blocks
   * returns). Output stays under ~30k characters.
   */
  readDocument(options: ReadOptions = {}): string {
    const blocks = this.view()
    const header = `The document has ${blocks.length} blocks (index|type|content preview):`
    // headings keep the wider preview even in tight mode (like the embedded agent)
    const render = (bodyMax: number) =>
      blocks.map((block, index) => {
        const preview = clip(
          blockPreviewText(block),
          block.type === 'heading' ? PREVIEW_MAX_CHARS : bodyMax,
        )
        return `${index}|${blockTypeName(block)}|${preview}`
      })
    const stats = (() => {
      const meta = this.meta()
      return `Full-text stats: words ${meta.wordCount}, characters (no spaces) ${meta.charCountNoSpaces}, characters (with spaces) ${meta.charCount}`
    })()

    const selected = this.selectedIndexes(options)
    if (selected === null) {
      let body = render(PREVIEW_MAX_CHARS)
      let out = [header, ...body, stats].join('\n')
      if (out.length > CONTEXT_MAX_CHARS) {
        // tighten previews first, then elide the middle (numbering stays verifiable)
        body = render(PREVIEW_TIGHT_CHARS)
        out = [header, ...body, stats].join('\n')
        if (out.length > CONTEXT_MAX_CHARS) {
          const dropStart = Math.floor(body.length / 3)
          const dropEnd = body.length - Math.floor(body.length / 3)
          out = [
            header,
            ...body.slice(0, dropStart),
            `…(${dropEnd - dropStart} blocks elided here; numbering is continuous)…`,
            ...body.slice(dropEnd),
            stats,
          ].join('\n')
        }
      }
      return out
    }

    const html = blocksToHtml(selected.map((i) => blocks[i]!))
    const clipped =
      html.length > READ_MAX_CHARS
        ? `${html.slice(0, READ_MAX_CHARS)}\n…(output truncated at ${READ_MAX_CHARS} characters; request a narrower block range)`
        : html
    return [
      header,
      ...render(PREVIEW_MAX_CHARS),
      stats,
      '',
      'Selected block content (restricted HTML):',
      clipped,
    ].join('\n')
  }

  private selectedIndexes(options: ReadOptions): number[] | null {
    const count = this.entries.length
    let indexes: number[] | null = null
    if (options.blocks !== undefined) {
      indexes = options.blocks
    } else if (options.range !== undefined) {
      const { start, end } = options.range
      indexes = []
      for (let i = start; i <= end; i++) indexes.push(i)
    }
    if (indexes === null) return null
    const valid = indexes.filter((i) => Number.isInteger(i) && i >= 0 && i < count)
    const invalid = indexes.length - valid.length
    if (invalid > 0)
      throw new Error(
        `${invalid} of the requested block indexes are out of range (document has ${count} blocks)`,
      )
    if (valid.length === 0) throw new Error('No blocks selected (empty blocks/range)')
    return [...new Set(valid)].sort((a, b) => a - b)
  }

  // ---- editing ----

  /**
   * Insert restricted-HTML content after block `at` (-1 = document start).
   * Lists get a real numbering id allocated up front (appended to
   * word/numbering.xml on save); everything else maps onto GeneratedBlocks.
   */
  insertContent(html: string, at?: number): { inserted: number; at: number } {
    const parsedHtml = parseRestrictedHtml(html)
    if (parsedHtml.length === 0) throw new Error('html did not parse into any content blocks')
    const count = this.entries.length
    const after =
      at === undefined || at === null
        ? count - 1
        : Math.min(Math.max(Math.trunc(at), -1), count - 1)
    const newEntries: SessionEntry[] = parsedHtml.map((block) => {
      if (block.type === 'table') return { kind: 'table', xml: block.xml, model: block.model }
      if (block.type === 'listItem' && block.list) {
        const numId = allocateNumId(block.list.kind, this.parsed, this.pending)
        return {
          kind: 'edited',
          original: null,
          gen: { ...block, list: { ...block.list, numId } },
        }
      }
      return { kind: 'edited', original: null, gen: block }
    })
    const insertAt = after + 1
    this.entries.splice(insertAt, 0, ...newEntries)
    return { inserted: newEntries.length, at: after }
  }

  /** Apply a batch of canonical ops atomically (validate forward, all or nothing). */
  applyOps(ops: Op[], dryRun = false): { results: OpResult[]; summary: string; dryRun: boolean } {
    const outcome = executeOps(this.entries, ops, this.parsed)
    if (!outcome.ok || !outcome.entries || !outcome.numbering) {
      throw new Error(
        `${outcome.error ?? 'ops failed'} — nothing was applied (atomic); fix and resend the whole batch`,
      )
    }
    if (!dryRun) {
      this.entries = outcome.entries
      this.pending.newDefs.push(...outcome.numbering.newDefs)
      this.pending.restartNums.push(...outcome.numbering.restartNums)
    }
    return {
      results: outcome.results,
      summary: outcome.results
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
   * Save atomically: the engine serializes the document (byte-preserving for
   * untouched blocks; with zero edits the original bytes round-trip verbatim),
   * then the bytes land via tmp-file + rename inside the target directory.
   *
   * mtime/size fencing: saving over the file this session opened refuses when
   * the file changed on disk since open (external writer), with a clear error.
   */
  async save(rawPath?: string): Promise<SaveResult> {
    const target = resolveConfined(rawPath ?? this.path, this.root)
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

    const saveBlocks = toSaveBlocks(this.entries)
    // only numbering definitions the final content actually references are
    // appended (an allocated-but-unused id — e.g. a setList that matched nothing
    // — must not touch word/numbering.xml)
    const referenced = new Set(
      this.entries
        .filter((e): e is Extract<SessionEntry, { kind: 'edited' }> => e.kind === 'edited')
        .map((e) => e.gen.list?.numId)
        .filter((id): id is string => id !== undefined),
    )
    const numbering = {
      newDefs: this.pending.newDefs.filter((d) => referenced.has(d.numId)),
      restartNums: this.pending.restartNums.filter((r) => referenced.has(r.numId)),
    }
    const hasNumbering = numbering.newDefs.length > 0 || numbering.restartNums.length > 0
    const bytes = await saveDocx(this.parsed, saveBlocks as SaveBlock[], {
      ...(hasNumbering ? { numbering } : {}),
    })

    await mkdir(dirname(target), { recursive: true })
    const tmp = join(dirname(target), `.${target.split('/').pop() ?? 'doc'}.airy-${randomUUID()}`)
    await writeFile(tmp, bytes)
    await rename(tmp, target)

    // refresh the fence so chained saves keep working
    if (target === this.path) {
      try {
        const info = await stat(this.path)
        this.baseline = { mtimeMs: info.mtimeMs, size: info.size }
      } catch {
        this.baseline = null
      }
    }
    this.savedPath = target
    return {
      path: target,
      bytes: bytes.byteLength,
      unchanged: bytes === this.parsed.internal.originalBytes,
    }
  }
}

// ---- session store ----

const sessions = new Map<string, DocxSession>()

export function storeSession(session: DocxSession): DocxSession {
  sessions.set(session.handle, session)
  return session
}

export function getSession(handle: string): DocxSession {
  const session = sessions.get(handle)
  if (!session)
    throw new Error(
      `Unknown document handle "${handle}". Open the document first with open_document.`,
    )
  return session
}

export function sessionCount(): number {
  return sessions.size
}

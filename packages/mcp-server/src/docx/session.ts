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
import {
  copyFile,
  link,
  mkdir,
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join, resolve } from 'node:path'

import {
  parseDocx,
  saveDocx,
  type Block,
  type GeneratedBlock,
  type ParsedDocFull,
  type SaveBlock,
} from '@airy-office/docx-engine'

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
import {
  convertViaSoffice,
  findSoffice,
  SOFFICE_FILTERS,
  sofficeMissingError,
} from '../import/soffice.js'

// ---- limits (mirror the embedded agent, scaled to the MCP 30k answer budget) ----

const CONTEXT_MAX_CHARS = 30_000
const PREVIEW_MAX_CHARS = 60
const PREVIEW_TIGHT_CHARS = 20
const READ_MAX_CHARS = 30_000
/**
 * Largest range span a read materializes: the request schema does not bound
 * `end`, so the session must reject a huge span BEFORE building the index
 * array (a range like 0..2^53 would otherwise hang/OOM the server).
 */
const RANGE_MAX_SPAN = 10_000

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
  kind: 'docx'
  path: string
  fileName: string
  format: 'docx'
  converted: boolean
  editable: boolean
  warnings: string[]
  originPath?: string
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
  /** which writer produced the file: the docx engine or a LibreOffice origin export */
  format?: 'docx' | 'origin'
  warnings?: string[]
}

/**
 * A legacy/ODF document this session was converted from (Phase 3b): editing
 * happens on a temp .docx produced by LibreOffice, `origin` remembers where
 * to write back on save_document(format:'origin').
 */
export interface SessionOrigin {
  /** absolute path of the original .doc/.odt file */
  readonly path: string
  readonly format: 'doc' | 'odt'
  /** mtime/size of the origin at conversion time (fence for origin export) */
  readonly stamp: { mtimeMs: number; size: number } | null
  /** temp dir holding the converted .docx; removed on close */
  readonly tempDir: string | null
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

/** the actionable clobber error shared by the guard and the exclusive promote */
export function saveTargetExistsError(target: string): Error {
  return new Error(
    `Refusing to save: "${target}" already exists and is not a file this session opened or ` +
      'saved. Pass overwrite: true to replace it (the existing file will be lost).',
  )
}

/**
 * Save-as clobber guard shared by the docx and xlsx sessions: refuse a target
 * that already exists on disk unless it is one of the session's own files
 * (the opened/backing file, a previous save's output) or the caller passed
 * overwrite. Default targets are guarded by the same rule: a native session
 * defaults to the file it opened (owned, so repeat saves keep working),
 * while a session converted from .xls/.ods/.doc/.odt defaults to a FRESH
 * sibling the session neither opened nor saved — a pre-existing sibling must
 * not be clobbered without consent.
 */
export async function assertSaveTargetFree(
  target: string,
  owned: ReadonlyArray<string | null>,
  overwrite: boolean | undefined,
): Promise<void> {
  if (overwrite === true) return
  if (owned.some((path) => path !== null && path === target)) return
  let exists = true
  try {
    await stat(target)
  } catch {
    exists = false
  }
  if (exists) throw saveTargetExistsError(target)
}

/**
 * Promote a temp file onto a guarded (fresh) save-as target WITHOUT replacing
 * an existing file: fs.link fails with EEXIST atomically, closing the window
 * between assertSaveTargetFree's stat and the write (TOCTOU) — a file
 * another writer created in that window surfaces the same actionable
 * overwrite error instead of being silently replaced by a rename. Saves that
 * replace by intent (same-file, the session's own output, overwrite:true)
 * keep using rename.
 *
 * exFAT/FAT/network shares do not support hard links, so link fails there
 * with EPERM/EACCES: on those codes the target is stat-checked and, when
 * still missing, promoted with the plain atomic rename instead. Trade-off
 * (accepted, link-less volumes only): rename REPLACES an existing file, so
 * the exclusive-create guarantee narrows to that stat — a file another
 * writer creates between the stat and the rename would be clobbered.
 * Volumes that do support links keep the race-free guarantee.
 */
export async function promoteNewFileExclusively(tmp: string, target: string): Promise<void> {
  try {
    await link(tmp, target)
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code
    if (code === 'EEXIST') {
      await rm(tmp, { force: true })
      throw saveTargetExistsError(target)
    }
    if (code === 'EPERM' || code === 'EACCES') {
      const targetExists = await stat(target).then(
        () => true,
        (statError: NodeJS.ErrnoException) => {
          if (statError.code === 'ENOENT') return false
          throw statError
        },
      )
      if (targetExists) {
        await rm(tmp, { force: true })
        throw saveTargetExistsError(target)
      }
      // the docx save's own atomic promote (rename consumes the temp file)
      await rename(tmp, target)
      return
    }
    await rm(tmp, { force: true })
    throw e
  }
  await rm(tmp, { force: true })
}

interface FileStamp {
  mtimeMs: number
  size: number
}

export class DocxSession {
  readonly handle: string
  /** absolute path the document was opened from (temp .docx for imports) */
  readonly path: string
  /** conversion origin when the document came from .doc/.odt */
  readonly origin: SessionOrigin | null

  private readonly root: string
  private readonly parsed: ParsedDocFull
  private entries: SessionEntry[]
  private pending = emptyNumbering()
  private baseline: FileStamp | null
  private savedPath: string | null = null
  /** every target this session has written (repeat save-as needs no overwrite) */
  private readonly savedTargets = new Set<string>()

  private constructor(
    handle: string,
    path: string,
    root: string,
    parsed: ParsedDocFull,
    stamp: FileStamp | null,
    origin: SessionOrigin | null,
  ) {
    this.handle = handle
    this.path = path
    this.root = root
    this.parsed = parsed
    this.entries = DocxSession.initialEntries(parsed)
    this.baseline = stamp
    this.origin = origin
  }

  private static initialEntries(parsed: ParsedDocFull): SessionEntry[] {
    return parsed.blocks
      .filter((block) => !block.hidden)
      .map((block) => ({ kind: 'original' as const, block }))
  }

  /**
   * Open a .docx inside the workspace root and parse it into the session
   * model. With `origin`, the path is a server-controlled temp .docx from a
   * .doc/.odt conversion (outside the root by design), so confinement is
   * skipped here — the origin path was confined before conversion.
   */
  static async open(rawPath: string, root?: string, origin?: SessionOrigin): Promise<DocxSession> {
    const path = origin ? resolve(rawPath) : resolveConfined(rawPath, root)
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
    return new DocxSession(
      randomUUID(),
      path,
      root ?? workspaceRoot(),
      parsed,
      stamp,
      origin ?? null,
    )
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
      kind: 'docx',
      path: this.origin?.path ?? this.path,
      fileName: (this.origin?.path ?? this.path).split('/').pop() ?? this.path,
      format: 'docx',
      converted: this.origin !== null,
      editable: true,
      warnings:
        this.origin === null
          ? []
          : [
              `Converted from .${this.origin.format} via LibreOffice — edits are applied to the converted .docx model.`,
            ],
      ...(this.origin === null ? {} : { originPath: this.origin.path }),
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
    if (options.blocks !== undefined) {
      const valid = options.blocks.filter((i) => Number.isInteger(i) && i >= 0 && i < count)
      const invalid = options.blocks.length - valid.length
      if (invalid > 0)
        throw new Error(
          `${invalid} of the requested block indexes are out of range (document has ${count} blocks)`,
        )
      if (valid.length === 0) throw new Error('No blocks selected (empty blocks/range)')
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
          `range ${start}..${end} spans ${span} blocks; the cap is ${RANGE_MAX_SPAN} per read ` +
            '(split large ranges into smaller reads)',
        )
      }
      const validEnd = Math.min(end, count - 1)
      const validCount = Math.max(0, validEnd - start + 1)
      const invalid = span - validCount
      if (invalid > 0)
        throw new Error(
          `${invalid} of the requested block indexes are out of range (document has ${count} blocks)`,
        )
      if (validCount === 0) throw new Error('No blocks selected (empty blocks/range)')
      const indexes: number[] = []
      for (let i = start; i <= validEnd; i++) indexes.push(i)
      return indexes
    }
    return null
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
   *
   * Save-as clobber guard: a target (explicit or default) that already exists
   * on disk is refused unless it is the file this session opened (or last
   * saved) or `overwrite` is true — a converted session's fresh sibling
   * default target is guarded like any save-as, while the native default
   * (the opened file) keeps working unchanged.
   *
   * Default target: the opened .docx; for sessions converted from .doc/.odt a
   * fresh sibling .docx next to the original. format:'origin' exports the
   * edited document back to the original .doc/.odt through LibreOffice
   * (best-effort) instead.
   */
  async save(
    rawPath?: string,
    format: 'docx' | 'origin' = 'docx',
    options: { overwrite?: boolean } = {},
  ): Promise<SaveResult> {
    if (format === 'origin') return this.saveToOrigin()
    const target = resolveConfined(rawPath ?? this.defaultTarget(), this.root)
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

    const bytes = await this.serialize()

    await mkdir(dirname(target), { recursive: true })
    const tmp = join(dirname(target), `.${target.split('/').pop() ?? 'doc'}.airy-${randomUUID()}`)
    await writeFile(tmp, bytes)
    // a fresh (guarded) target promotes exclusively: a file created between
    // the guard's stat and this write surfaces the clobber error instead of
    // being silently replaced; targets this session owns replace by intent
    if (options.overwrite === true || target === this.path || this.savedTargets.has(target)) {
      await rename(tmp, target)
    } else {
      await promoteNewFileExclusively(tmp, target)
    }

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
    this.savedTargets.add(target)
    return {
      path: target,
      bytes: bytes.byteLength,
      unchanged: bytes === this.parsed.internal.originalBytes,
      format: 'docx',
      ...(this.origin === null
        ? {}
        : {
            warnings: [
              `Saved as .docx; the original .${this.origin.format} file "${this.origin.path}" was left untouched (use format "origin" to export back).`,
            ],
          }),
    }
  }

  private defaultTarget(): string {
    if (this.origin === null) return this.path
    const name = basename(this.origin.path)
    const dot = name.lastIndexOf('.')
    const stem = dot > 0 ? name.slice(0, dot) : name
    return join(dirname(this.origin.path), `${stem}.docx`)
  }

  /** Engine serialization shared by the docx and origin save paths. */
  private async serialize(): Promise<Uint8Array> {
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
    return saveDocx(this.parsed, saveBlocks as SaveBlock[], {
      ...(hasNumbering ? { numbering } : {}),
    })
  }

  /**
   * Best-effort export back to the original legacy/ODF format: serialize the
   * edited docx in memory, convert through LibreOffice with the canonical
   * export filter, promote atomically onto the origin path.
   */
  private async saveToOrigin(): Promise<SaveResult> {
    if (this.origin === null) {
      throw new Error('format "origin" is only valid for sessions converted from .doc/.odt.')
    }
    if (this.origin.stamp) {
      let current: FileStamp | null
      try {
        const info = await stat(this.origin.path)
        current = { mtimeMs: info.mtimeMs, size: info.size }
      } catch {
        current = null
      }
      if (
        !current ||
        current.mtimeMs !== this.origin.stamp.mtimeMs ||
        current.size !== this.origin.stamp.size
      ) {
        throw new FencingError(this.origin.path)
      }
    }
    const tool = await findSoffice()
    if (!tool) {
      throw sofficeMissingError(`Exporting back to .${this.origin.format} requires LibreOffice.`)
    }
    const bytes = await this.serialize()
    const workDir = await mkdtemp(join(tmpdir(), 'airy-origin-'))
    try {
      const tempDocx = join(workDir, 'document.docx')
      await writeFile(tempDocx, bytes)
      const output = await convertViaSoffice(tool, tempDocx, {
        filter: this.origin.format === 'doc' ? SOFFICE_FILTERS.doc : SOFFICE_FILTERS.odt,
        extension: this.origin.format,
        outDir: workDir,
      })
      const tmpTarget = join(
        dirname(this.origin.path),
        `.${basename(this.origin.path)}.airy-${randomUUID()}`,
      )
      await copyFile(output, tmpTarget)
      await rename(tmpTarget, this.origin.path)
      this.savedPath = this.origin.path
      this.savedTargets.add(this.origin.path)
      const info = await stat(this.origin.path)
      return {
        path: this.origin.path,
        bytes: info.size,
        unchanged: false,
        format: 'origin',
        warnings: [
          `Exported back to .${this.origin.format} via LibreOffice — best-effort fidelity, verify the result.`,
        ],
      }
    } finally {
      await rm(workDir, { recursive: true, force: true })
    }
  }

  /** Release session resources: remove the conversion temp dir when present. */
  async close(): Promise<string[]> {
    if (this.origin?.tempDir) {
      await rm(this.origin.tempDir, { recursive: true, force: true })
      return [this.origin.tempDir]
    }
    return []
  }
}

// Session storage lives in ../sessions/store.ts (one handle space shared by
// the docx, xlsx and read-only text sessions since S6).

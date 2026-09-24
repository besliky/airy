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
import { classifyOleContent, encryptedOfficeRefusal } from '../import/ole.js'
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
import { resolveConfined, assertWorkspaceRootExists, workspaceRoot } from './paths.js'
import {
  convertViaSoffice,
  findSoffice,
  SOFFICE_FILTERS,
  sofficeMissingError,
} from '../import/soffice.js'
import { assertWithinOpenCap, OpenSizeError } from '../sessions/size-fence.js'

// ---- limits (mirror the embedded agent, scaled to the MCP 30k answer budget) ----

const PREVIEW_MAX_CHARS = 60
const PREVIEW_TIGHT_CHARS = 20
/**
 * Read-answer budget in UTF-8 bytes — the honest, documented limit (BUG-1684:
 * the budget is byte-shaped for the host that frames the answer, and char
 * counting underestimates CJK text threefold; the audit measured the old
 * answers at 800KB..2.9MB against this 30k budget).
 */
const READ_BUDGET_BYTES = 30_000
/**
 * Overview budget a SELECTION read (blocks/range) spends on the index
 * listing: the payload of such a read is the selected blocks' HTML, so the
 * overview is capped well below the full budget and the HTML takes the rest
 * (BUG-1684 — the selection path used to append the full overview of ALL
 * blocks, so a five-block range into a 42k-block document answered with
 * ~2.9MB against the 30k budget).
 */
const SELECTION_OVERVIEW_MAX_BYTES = 4_000
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

/** UTF-8 byte length — the read-answer budget is measured in bytes */
function byteLen(text: string): number {
  return Buffer.byteLength(text, 'utf8')
}

/**
 * Longest char-prefix of `text` whose UTF-8 bytes stay within `maxBytes`
 * (binary search over prefix length — prefix byte length is monotonic).
 * `reserve` holds room for a truncation note appended after the prefix.
 */
function bytePrefixLimit(text: string, maxBytes: number, reserve: number): number {
  const allowed = Math.max(0, maxBytes - reserve)
  let lo = 0
  let hi = text.length
  while (lo < hi) {
    const mid = Math.ceil((lo + hi + 1) / 2)
    if (byteLen(text.slice(0, mid)) <= allowed) lo = mid
    else hi = mid - 1
  }
  return lo
}

/**
 * Fit the overview (header + `index|type|preview` lines + stats) into
 * `budget` bytes: full previews first, then tightened previews, then
 * repeated middle elision — each round keeps half the lines (head from the
 * start, tail from the end, one marker naming the elided count) until the
 * output fits. The single-pass elide this loop replaces still left 2/3 of
 * the lines in, so a 42k-block document answered a default read with ~800k
 * characters against the 30k budget (BUG-1684). Termination: the kept count
 * halves each round down to a floor of two lines, where the assembled
 * overview is a few hundred bytes — always inside any budget the callers
 * pass.
 */
function fitOverview(
  header: string,
  stats: string,
  render: (bodyMax: number) => string[],
  budget: number,
): string {
  let body = render(PREVIEW_MAX_CHARS)
  let out = [header, ...body, stats].join('\n')
  if (byteLen(out) > budget) {
    // tighten previews first, then elide the middle (numbering stays verifiable)
    body = render(PREVIEW_TIGHT_CHARS)
    out = [header, ...body, stats].join('\n')
    let keep = body.length
    while (byteLen(out) > budget && keep > 2) {
      keep = Math.max(2, Math.floor(keep / 2))
      const head = Math.ceil(keep / 2)
      const tail = keep - head
      out = [
        header,
        ...body.slice(0, head),
        `…(${body.length - keep} blocks elided here; numbering is continuous)…`,
        ...body.slice(body.length - tail),
        stats,
      ].join('\n')
    }
  }
  return out
}

/**
 * Final guard for an assembled read answer, in the shape the line sessions
 * use (line-core clip): an honest hard cap with a truncation note. The
 * callers size their parts to fit by construction; this only bounds a
 * pathological part so no read answer can ever exceed the budget.
 */
function finalClip(text: string, hint: string): string {
  if (byteLen(text) <= READ_BUDGET_BYTES) return text
  const note = `\n…(output truncated; the read budget is ${READ_BUDGET_BYTES} bytes — ${hint})`
  return `${text.slice(0, bytePrefixLimit(text, READ_BUDGET_BYTES, byteLen(note)))}${note}`
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

/**
 * Run a save's tmp-write + promote with the orphan cleanup the individual
 * steps lack: when anything between creating the `.<name>.airy-<uuid>`
 * dotfile and its promote throws (a failed write, a crashed engine stream,
 * a rejected rename), the dotfile is removed instead of leaking next to the
 * target forever (BUG-1111). Best-effort — a cleanup failure never masks
 * the original save error. On success the promote consumed the tmp file
 * (rename) or removed it (link), so no cleanup runs at all.
 */
export async function withSaveTmpCleanup<T>(tmp: string, body: () => Promise<T>): Promise<T> {
  try {
    return await body()
  } catch (e) {
    await rm(tmp, { force: true }).catch(() => undefined)
    throw e
  }
}

/**
 * Permission-shaped failure of a save's file write (UX-1690): the raw error
 * ("EACCES: permission denied, open '<...>.airy-<uuid>'") names a temp
 * dotfile the agent never created and says nothing about the cause. Name the
 * actual cause instead: the target directory or the target file refuses
 * writes — a read-only workspace, file, or volume.
 */
export class WorkspaceWriteError extends Error {
  constructor(target: string, cause: unknown) {
    super(
      `Cannot write "${dirname(target)}": permission denied — the workspace directory or the ` +
        'target file may be read-only. Fix the permissions or save to another path.',
      { cause },
    )
    this.name = 'WorkspaceWriteError'
  }
}

/** errno codes that mean the save could not write: denied or read-only volume */
const WRITE_PERMISSION_CODES = new Set(['EACCES', 'EROFS'])

/**
 * True when a save failure is a permission refusal. The xlsx save writes
 * through the Rust sidecar, so its refusal arrives as a plain error without
 * a `.code` — match the errno text the sidecar relays too ("Permission
 * denied (os error 13)", "Read-only file system (os error 30)").
 */
export function isWritePermissionError(e: unknown): boolean {
  const code = (e as NodeJS.ErrnoException | null)?.code
  if (typeof code === 'string' && WRITE_PERMISSION_CODES.has(code)) return true
  return (
    e instanceof Error &&
    /EACCES|EROFS|permission denied|read-only file system|os error 13|os error 30/i.test(e.message)
  )
}

/**
 * Map a save-path failure to a WorkspaceWriteError when it is
 * permission-shaped (UX-1690); every other error passes through unchanged so
 * the established save contracts (fences, clobber guards, drift refusals)
 * keep their exact messages.
 */
export function asWorkspaceWriteError(e: unknown, target: string): unknown {
  return isWritePermissionError(e) ? new WorkspaceWriteError(target, e) : e
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
      // SEC-1102: refuse a runaway raw file before reading it into memory
      // (the zip-bomb budget itself runs inside parseDocx — docx-engine's
      // assertZipWithinLimits — so declared-uncompressed bombs are refused
      // before any entry is inflated)
      const info = await stat(path)
      assertWithinOpenCap(path, info.size, 'docx')
      bytes = new Uint8Array(await readFile(path))
      assertWithinOpenCap(path, bytes.byteLength, 'docx')
      stamp = { mtimeMs: info.mtimeMs, size: info.size }
    } catch (e) {
      if (e instanceof OpenSizeError) throw e
      throw new Error(`Cannot read "${path}": ${e instanceof Error ? e.message : String(e)}`, {
        cause: e,
      })
    }
    // BUG-1504: an encrypted .docx is not a zip — Word repackages it as an
    // OLE2 (CFB) container, so jszip's parse failure ("Can't find end of
    // central directory") reads like a corrupt file. Name password
    // protection explicitly; a plain OLE container is a renamed .doc.
    const ole = classifyOleContent(bytes)
    if (ole === 'encrypted-ooxml' || ole === 'encrypted-legacy') {
      throw encryptedOfficeRefusal(path, 'docx')
    }
    if (ole === 'plain') {
      throw new Error(
        `Cannot open "${path}": the file is an OLE2 compound document (a legacy .doc or other ` +
          'non-zip container), not a .docx package. Open it as .doc — LibreOffice converts it — ' +
          'or re-save it as .docx.',
      )
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
      // basename, not a '/'-split: on Windows a `\`-separated path never
      // split on '/', which made the label the whole path (BUG-706)
      fileName: basename(this.origin?.path ?? this.path) || this.path,
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
   * blocks is appended after an overview fitted into a bounded slice (the
   * same tighten+elide ladder). The whole answer stays under the 30k-byte
   * budget on both paths (BUG-1684).
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
      return finalClip(
        fitOverview(header, stats, render, READ_BUDGET_BYTES),
        'request specific blocks/range',
      )
    }

    const html = blocksToHtml(selected.map((i) => blocks[i]!))
    const label = 'Selected block content (restricted HTML):'
    // the overview rides along in a bounded slice so it can never crowd out
    // the selected content; the HTML payload takes the rest of the budget.
    // join('\n') below adds three separators between the four parts.
    const overview = fitOverview(header, stats, render, SELECTION_OVERVIEW_MAX_BYTES)
    const note = `\n…(output truncated; the read budget is ${READ_BUDGET_BYTES} bytes — request a narrower block range)`
    const htmlBudget = Math.max(
      1_000,
      READ_BUDGET_BYTES - byteLen(overview) - byteLen(label) - 3 - byteLen(note),
    )
    const clipped =
      byteLen(html) <= htmlBudget
        ? html
        : `${html.slice(0, bytePrefixLimit(html, htmlBudget, 0))}${note}`
    return finalClip([overview, '', label, clipped].join('\n'), 'request a narrower block range')
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
        .map((r) => {
          const base = `${r.op}: matched ${r.matched}, changed ${r.changed}${r.detail ? ` (${r.detail})` : ''}`
          // warnings ride the summary too: the text answer is what most
          // agents quote, so a compromise must never hide in the JSON only
          return r.warnings?.length ? `${base} — warning: ${r.warnings.join(' ')}` : base
        })
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
   * Stale-root refusal: when the pinned workspace root has been moved/renamed
   * since open, the save fails with StaleWorkspaceRootError instead of
   * re-creating the dead directory (the mkdir never runs).
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
    // a pinned root that vanished (moved/renamed workspace directory) must
    // fail here, before confinement + mkdir silently resurrect it (BUG-1103)
    await assertWorkspaceRootExists(this.root)
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
    // basename, not a '/'-split: on Windows the split leaves the whole path
    // in the temp name and writeFile fails on the colons/backslashes
    const tmp = join(dirname(target), `.${basename(target) || 'doc'}.airy-${randomUUID()}`)
    try {
      await withSaveTmpCleanup(tmp, async () => {
        await writeFile(tmp, bytes)
        // a fresh (guarded) target promotes exclusively: a file created between
        // the guard's stat and this write surfaces the clobber error instead of
        // being silently replaced; targets this session owns replace by intent
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
      // local capture: the null guard above does not reach inside the cleanup
      // closure, and the promote must target exactly the checked origin path
      const originPath = this.origin.path
      const tmpTarget = join(dirname(originPath), `.${basename(originPath)}.airy-${randomUUID()}`)
      try {
        await withSaveTmpCleanup(tmpTarget, async () => {
          await copyFile(output, tmpTarget)
          await rename(tmpTarget, originPath)
        })
      } catch (e) {
        // UX-1690: same friendly mapping as the docx save — the origin export
        // fails the same way in a read-only workspace
        throw asWorkspaceWriteError(e, originPath)
      }
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

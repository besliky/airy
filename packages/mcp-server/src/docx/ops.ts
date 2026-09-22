// Canonical edit ops for headless docx sessions — the engine-patch twin of the
// embedded agent's ops.ts. The model and the tools issue the same flat
// `{ op, target, ...fields }` records; a registry validates the whole batch up
// front and applies it as one atomic step (all or nothing), so a failed batch
// never leaves the session half-edited. Fields are patches: a present key sets
// the property, null clears it, an absent key leaves it untouched.
//
// Ops are applied directly to the docx-engine Block model: an edited original
// block is converted into a GeneratedBlock (keeping its docxIndex anchor
// semantics intact via save-time rawPPr passthrough), so untouched blocks stay
// byte-identical on save.
import {
  mergePPrFormat,
  type Block,
  type GeneratedBlock,
  type ParaFormat,
  type Run,
  type TableModel,
} from '@airy-office/docx-engine'

import { caseInsensitiveRanges, replaceCaseInsensitive, type FoldRange } from '../case-fold.js'

// ---- session entry model ----

/**
 * One addressable block in a session. 'original' entries reference the parsed
 * engine Block (never mutated — that is the byte-preservation guarantee);
 * 'edited' entries carry a GeneratedBlock rebuilt from an original (or fully
 * new); 'table' entries are self-contained w:tbl fragments with the model kept
 * for read-back serialization.
 */
export type SessionEntry =
  | { kind: 'original'; block: Block }
  | { kind: 'edited'; original: Block | null; gen: GeneratedBlock }
  | { kind: 'table'; xml: string; model: TableModel }

/** numbering definitions that must be appended to word/numbering.xml on save */
export interface PendingNumbering {
  newDefs: Array<{ numId: string; kind: 'bullet' | 'ordered' }>
  restartNums: Array<{
    numId: string
    abstractNumId: string
    startOverrides: Record<number, number>
  }>
}

export function emptyNumbering(): PendingNumbering {
  return { newDefs: [], restartNums: [] }
}

/** the numbering map shape the engine's ParsedDoc exposes (allocation source) */
export interface NumberingSource {
  numbering: Map<string, { abstractNumId: string; levels: Record<number, { numFmt: string }> }>
  /**
   * heading level -> styleId present in the document (ParsedDoc exposes it).
   * setHeadingLevel reads it to report when the target level has no style, so
   * the op can disclose the serializer's outline-level fallback (BUG-1631)
   * instead of a bare "changed N".
   */
  headingStyleIds?: ReadonlyMap<number, string>
}

// ---- op / target contracts (mirrors the embedded agent's ops.ts) ----

export interface Target {
  /** canonical spelling here; the renderer aliases (docHeading/docParagraph/docListItem) normalize onto them */
  nodeType?:
    'heading' | 'paragraph' | 'listItem' | 'image' | 'docHeading' | 'docParagraph' | 'docListItem'
  headingLevel?: number
  containsText?: string
  matchCase?: boolean
  blockIndexes?: number[]
}

export interface Op {
  op: string
  target?: Target
  [key: string]: unknown
}

export interface OpResult {
  op: string
  matched: number
  changed: number
  skippedProtected: number
  detail?: string
  /**
   * Non-fatal compromises the op made, when any. "changed" blocks survive
   * save regardless; the warnings explain HOW (BUG-1631: a heading level
   * with no style in the document is applied as a direct outline-level
   * override and the paragraph keeps its previous formatting). Optional —
   * absent when the op had nothing to disclose, so the wire shape of the
   * common path is unchanged.
   */
  warnings?: string[]
}

export interface ExecuteOutcome {
  ok: boolean
  results: OpResult[]
  error?: string
}

const FONT_KEYS = [
  'bold',
  'italic',
  'underline',
  'strike',
  'color',
  'highlight',
  'fontSize',
  'fontFamily',
  'baseline',
] as const
const PARA_KEYS = [
  'align',
  'lineSpacing',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'shadingFill',
  'borders',
] as const
const NODE_TYPES = ['heading', 'paragraph', 'listItem', 'image'] as const
/** the renderer's canonical spellings, accepted as aliases (both directions documented in docs/COPILOT.md) */
const NODE_TYPE_ALIASES = new Map([
  ['docHeading', 'heading'],
  ['docParagraph', 'paragraph'],
  ['docListItem', 'listItem'],
])
/** all spellings the validator accepts, for error messages */
const NODE_TYPE_SPELLINGS = [
  'heading|docHeading',
  'paragraph|docParagraph',
  'listItem|docListItem',
  'image',
] as const
const BASELINES = ['superscript', 'subscript', 'none'] as const
const ALIGNS = ['left', 'center', 'right', 'justify'] as const

const HEX = /^#?([0-9a-f]{6})$/i

/** "#RRGGBB" or "RRGGBB" -> "RRGGBB"; null passes through */
function normalizeHex(value: unknown, where: string, field: string): string | null {
  if (value === null) return null
  const m = typeof value === 'string' ? HEX.exec(value.trim()) : null
  if (!m) {
    throw new Error(
      `${where}: ${field} must be a 6-digit hex color like "#1A73E8" (or null to clear)`,
    )
  }
  return m[1].toUpperCase()
}

function checkHex(value: unknown, where: string, field: string): string | null {
  try {
    normalizeHex(value, where, field)
    return null
  } catch (e) {
    return e instanceof Error ? e.message : String(e)
  }
}

const present = (op: Op, keys: readonly string[]): string[] =>
  keys.filter((k) => op[k] !== undefined)
const isNumberOrNull = (v: unknown) => v === null || (typeof v === 'number' && Number.isFinite(v))
const isPositiveOrNull = (v: unknown) =>
  v === null || (typeof v === 'number' && Number.isFinite(v) && v > 0)
const isStringOrNull = (v: unknown) => v === null || typeof v === 'string'

/** entry kinds the text ops may touch */
const TEXT_TYPES = new Set(['paragraph', 'heading', 'listItem'])

// ---- registry ----

interface OpDef {
  name: string
  signature: string
  keys: readonly string[]
  target: 'required' | 'optional' | 'none'
  validate(op: Op, where: string): string | null
  apply(op: Op, env: OpEnv): OpResult
}

interface OpEnv {
  entries: SessionEntry[]
  doc: NumberingSource
  numbering: PendingNumbering
}

const REGISTRY = new Map<string, OpDef>()

function register(def: OpDef): void {
  REGISTRY.set(def.name, def)
}

export function opNames(): string[] {
  return [...REGISTRY.values()].map((d) => d.name)
}

export function opSignatures(): string[] {
  return [...REGISTRY.values()].map((d) => d.signature)
}

// ---- target matching ----

type EntryType = 'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'passthrough'

/** effective type of an entry for nodeType matching */
function entryType(entry: SessionEntry): EntryType {
  if (entry.kind === 'original') return entry.block.type
  if (entry.kind === 'table') return 'table'
  return entry.gen.type
}

/** current live text of an entry (tracked deletions excluded, like liveText) */
function entryText(entry: SessionEntry): string {
  if (entry.kind === 'table') return ''
  if (entry.kind === 'original') {
    const b = entry.block
    if (b.type === 'paragraph' || b.type === 'heading' || b.type === 'listItem')
      return (b.runs ?? [])
        .filter((r) => !r.del)
        .map((r) => r.text)
        .join('')
    return [b.label, b.previewText].filter(Boolean).join(' ')
  }
  return (entry.gen.runs ?? [])
    .filter((r) => !r.del)
    .map((r) => r.text)
    .join('')
}

/**
 * Map the renderer's alias spellings onto this registry's canonical node
 * types — the mirror of the renderer-side normalizeNodeType
 * (apps/docs/src/renderer/ai/ops.ts), which maps this vocabulary onto the PM
 * node names. Both spellings therefore target the same blocks in both the
 * headless and the live registries; unknown spellings pass through unchanged
 * and are rejected by the validator.
 */
export function normalizeNodeType(nodeType: string): string {
  return NODE_TYPE_ALIASES.get(nodeType) ?? nodeType
}

function matchTarget(entries: SessionEntry[], target: Target): number[] {
  const out: number[] = []
  const nodeType = target.nodeType === undefined ? undefined : normalizeNodeType(target.nodeType)
  entries.forEach((entry, index) => {
    if (target.blockIndexes && !target.blockIndexes.includes(index)) return
    const type = entryType(entry)
    if (nodeType && type !== nodeType) return
    if (target.headingLevel !== undefined) {
      if (type !== 'heading') return
      const level =
        entry.kind === 'original'
          ? entry.block.level
          : entry.kind === 'edited'
            ? (entry.gen.level ?? 1)
            : undefined
      if (Number(level) !== target.headingLevel) return
    }
    if (target.containsText) {
      const text = entryText(entry)
      const hit =
        target.matchCase === false
          ? text.toLowerCase().includes(target.containsText.toLowerCase())
          : text.includes(target.containsText)
      if (!hit) return
    }
    out.push(index)
  })
  return out
}

// ---- editing primitives ----

/** deep copy a GeneratedBlock (plain JSON data, so a JSON round trip is exact) */
function cloneGen(gen: GeneratedBlock): GeneratedBlock {
  return JSON.parse(JSON.stringify(gen)) as GeneratedBlock
}

/** the edit target of editEntry */
interface EditHandle {
  gen: GeneratedBlock
  /**
   * call after the op's mutations: reverts the entry to its original state
   * when nothing actually changed, so no-op ops keep blocks byte-identical
   * (returns true when the block did change)
   */
  commit(): boolean
}

/**
 * Convert an 'original' entry in place into an 'edited' one (deep copy of the
 * model); an already-edited entry is returned as-is. Mutate `gen`, then call
 * commit() — a no-op mutation reverts the entry, preserving its bytes.
 */
function editEntry(entries: SessionEntry[], index: number): EditHandle {
  const entry = entries[index]!
  if (entry.kind === 'edited') {
    const before = JSON.stringify(entry.gen)
    const revertTo: SessionEntry = { ...entry, gen: JSON.parse(before) }
    return {
      gen: entry.gen,
      commit: () => {
        if (JSON.stringify(entry.gen) === before) {
          entries[index] = revertTo
          return false
        }
        return true
      },
    }
  }
  if (entry.kind === 'original') {
    const block = entry.block
    const gen: GeneratedBlock = {
      type:
        block.type === 'paragraph' || block.type === 'heading' || block.type === 'listItem'
          ? block.type
          : 'paragraph',
      ...(block.level !== undefined ? { level: block.level } : {}),
      ...(block.outlineOnly ? { outlineOnly: true } : {}),
      ...(block.styleId !== undefined ? { styleId: block.styleId } : {}),
      ...(block.list ? { list: { ...block.list } } : {}),
      ...(block.format ? (JSON.parse(JSON.stringify(block.format)) as ParaFormat) : {}),
      ...(block.bookmarks?.length ? { bookmarks: [...block.bookmarks] } : {}),
      ...(block.hiddenBookmarks?.length ? { hiddenBookmarks: [...block.hiddenBookmarks] } : {}),
      ...(block.commentStarts?.length ? { commentStarts: [...block.commentStarts] } : {}),
      ...(block.commentEnds?.length ? { commentEnds: [...block.commentEnds] } : {}),
      runs: (block.runs ?? []).map((run) => ({ ...run })),
      ...(block.sdtShell ? { sdtShell: { ...block.sdtShell } } : {}),
      ...(block.pPrChangeInfo ? { pPrChange: JSON.stringify(block.pPrChangeInfo) } : {}),
      ...(block.blockRevision ? { blockRevision: { ...block.blockRevision } } : {}),
    }
    entries[index] = { kind: 'edited', original: block, gen }
    const before = JSON.stringify(gen)
    return {
      gen,
      commit: () => {
        if (JSON.stringify(gen) === before) {
          entries[index] = { kind: 'original', block }
          return false
        }
        return true
      },
    }
  }
  throw new Error('Tables cannot be edited in place (delete and re-insert instead)')
}

function indentSame(a: ParaFormat | undefined, b: ParaFormat | undefined): boolean {
  const norm = (v: number | undefined) => (v !== undefined ? Math.round(v) : null)
  return (
    norm(a?.indentLeft) === norm(b?.indentLeft) &&
    norm(a?.indentRight) === norm(b?.indentRight) &&
    norm(a?.indentFirstLine) === norm(b?.indentFirstLine)
  )
}

function normalizedFormat(format: ParaFormat | undefined): unknown {
  if (!format) return null
  return [
    format.align ?? null,
    format.lineSpacing ?? null,
    format.lineRule ?? null,
    format.lineRawTwips ?? null,
    format.bidi ?? false,
    format.indentLeft ?? null,
    format.indentRight ?? null,
    format.indentFirstLine ?? null,
    format.spaceBefore ?? null,
    format.spaceAfter ?? null,
    format.pageBreakBefore ?? false,
    format.shadingFill ?? null,
    format.borders ?? null,
    format.borderLines ? JSON.stringify(format.borderLines) : null,
    format.tabStops ? JSON.stringify(format.tabStops) : null,
    format.dropCap ? JSON.stringify(format.dropCap) : null,
    format.autoSpace ?? null,
    format.wordWrap ?? null,
    format.overflowPunct ?? null,
    format.emptyRunSizeHalfPoints ?? null,
  ]
}

/**
 * High-fidelity pPr passthrough for in-place paragraph edits (port of the
 * renderer's applyRawPPr): text-only edits reuse the original <w:pPr> bytes
 * verbatim; format edits merge the modeled children into them; structure
 * changes fall back to a full rebuild.
 */
export function applyRawPPr(gen: GeneratedBlock, original: Block | null): void {
  if (!original) return
  const structureSame =
    original.type === gen.type &&
    (original.styleId ?? null) === (gen.styleId ?? null) &&
    JSON.stringify(original.list ?? null) === JSON.stringify(gen.list ?? null)
  if (original.rawPPr === undefined) {
    if (
      !original.format?.charIndents ||
      !structureSame ||
      gen.type !== 'paragraph' ||
      indentSame(original.format, gen.format)
    )
      return
    gen.rawPPr = mergePPrFormat('', gen.format, original.format)
    return
  }
  if (!structureSame) return
  const formatSame =
    JSON.stringify(normalizedFormat(original.format)) ===
    JSON.stringify(normalizedFormat(gen.format))
  gen.rawPPr = formatSame
    ? original.rawPPr
    : mergePPrFormat(original.rawPPr, gen.format, original.format)
}

/** apply font fields to one run copy (present=set, null=clear, absent=untouched) */
function applyFontToRun(run: Run, op: Op, where: string): void {
  if (op.bold !== undefined) run.bold = op.bold === true ? true : undefined
  if (op.italic !== undefined) run.italic = op.italic === true ? true : undefined
  if (op.underline !== undefined) run.underline = op.underline === true ? true : undefined
  if (op.strike !== undefined) run.strike = op.strike === true ? true : undefined
  if (op.color !== undefined) {
    const hex = normalizeHex(op.color, where, 'color')
    if (hex) run.color = hex
    else delete run.color
  }
  if (op.highlight !== undefined) {
    const v = op.highlight as string | null
    if (v) run.highlight = v
    else delete run.highlight
  }
  if (op.fontSize !== undefined) {
    if (op.fontSize === null) delete run.sizeHalfPoints
    else run.sizeHalfPoints = Math.round((op.fontSize as number) * 2)
  }
  if (op.fontFamily !== undefined) {
    const v = op.fontFamily as string | null
    if (v) {
      run.font = v
      run.fontAscii = v
    } else {
      delete run.font
      delete run.fontAscii
    }
  }
  if (op.baseline !== undefined) {
    const v = op.baseline
    if (v === 'superscript' || v === 'subscript') run.vertAlign = v
    else delete run.vertAlign
  }
}

/** split runs so each occurrence of `needle` becomes its own styled run; returns the hit count */
function styleOccurrences(
  runs: Run[],
  needle: string,
  matchCase: boolean,
  apply: (run: Run) => void,
): number {
  let count = 0
  const out: Run[] = []
  for (const run of runs) {
    if (needle === '') {
      out.push(run)
      continue
    }
    // match ranges in ORIGINAL run coordinates: the case-insensitive fold is
    // not length-preserving (İ expands), so lowered-index slices corrupted
    // runs with Turkish text (BUG-1101)
    const ranges = matchCase
      ? literalRanges(run.text, needle)
      : caseInsensitiveRanges(run.text, needle)
    if (ranges.length === 0) {
      out.push(run)
      continue
    }
    let pos = 0
    for (const { start, end } of ranges) {
      if (start > pos) out.push({ ...run, text: run.text.slice(pos, start) })
      const hit: Run = { ...run, text: run.text.slice(start, end) }
      apply(hit)
      out.push(hit)
      count++
      pos = end
    }
    if (pos < run.text.length) out.push({ ...run, text: run.text.slice(pos) })
  }
  runs.length = 0
  runs.push(...out)
  return count
}

/** all non-overlapping case-SENSITIVE literal matches as original-coordinate ranges */
function literalRanges(text: string, needle: string): FoldRange[] {
  const ranges: FoldRange[] = []
  let idx = text.indexOf(needle)
  while (idx !== -1) {
    ranges.push({ start: idx, end: idx + needle.length })
    idx = text.indexOf(needle, idx + needle.length)
  }
  return ranges
}

// ---- numbering allocation (renderer allocateListNumId, headless variant) ----

/** max numId across the document's numbering part and the pending additions */
function maxNumId(doc: NumberingSource, pending: PendingNumbering): number {
  let max = 2 // the blank template occupies 1/2
  for (const id of doc.numbering.keys()) max = Math.max(max, parseInt(id, 10) || 0)
  for (const d of pending.newDefs) max = Math.max(max, parseInt(d.numId, 10) || 0)
  for (const r of pending.restartNums) max = Math.max(max, parseInt(r.numId, 10) || 0)
  return max
}

/**
 * Allocate a numId for a list kind, reusing an existing same-kind abstractNum
 * (restart at 1) or creating a blank-template definition. Mutates `pending`.
 */
export function allocateNumId(
  kind: 'bullet' | 'ordered',
  doc: NumberingSource,
  pending: PendingNumbering,
): string {
  const match = [...doc.numbering.values()].find(
    (d) => (d.levels[0]?.numFmt === 'bullet') === (kind === 'bullet'),
  )
  const numId = String(maxNumId(doc, pending) + 1)
  if (match) {
    pending.restartNums.push({
      numId,
      abstractNumId: match.abstractNumId,
      startOverrides: { 0: 1 },
    })
  } else {
    pending.newDefs.push({ numId, kind })
  }
  return numId
}

// ---- op executors ----

function textIndexes(
  env: OpEnv,
  target: Target,
): { text: number[]; skipped: number; matched: number } {
  const matched = matchTarget(env.entries, target)
  const text: number[] = []
  let skipped = 0
  for (const index of matched) {
    if (TEXT_TYPES.has(entryType(env.entries[index]!))) text.push(index)
    else skipped++
  }
  return { text, skipped, matched: matched.length }
}

register({
  name: 'setFont',
  signature:
    'setFont <target> bold? italic? underline? strike? color? highlight? fontSize? fontFamily? baseline? — style whole blocks',
  keys: FONT_KEYS,
  target: 'required',
  validate: (op, where) => validateFontFields(op, where),
  apply: (op, env) => {
    const { text, skipped, matched } = textIndexes(env, op.target as Target)
    let changed = 0
    for (const index of text) {
      const edit = editEntry(env.entries, index)
      for (const run of edit.gen.runs ?? []) applyFontToRun(run, op, 'setFont')
      if (edit.commit()) changed++
    }
    return { op: op.op, matched, changed, skippedProtected: skipped }
  },
})

register({
  name: 'setMatchedFont',
  signature:
    'setMatchedFont text matchCase? <target?> + font fields — style only the occurrences of text',
  keys: ['text', 'matchCase', ...FONT_KEYS],
  target: 'optional',
  validate: (op, where) => {
    if (typeof op.text !== 'string' || op.text === '')
      return `${where}: text must be a non-empty string`
    return validateFontFields(op, where)
  },
  apply: (op, env) => {
    const needle = op.text as string
    const matchCase = op.matchCase !== false
    const { text, skipped, matched } = op.target
      ? textIndexes(env, op.target)
      : {
          text: allTextIndexes(env.entries),
          skipped: countProtected(env.entries),
          matched: env.entries.length,
        }
    let changed = 0
    let hits = 0
    for (const index of text) {
      const edit = editEntry(env.entries, index)
      const n = styleOccurrences(edit.gen.runs ?? [], needle, matchCase, (run) =>
        applyFontToRun(run, op, 'setMatchedFont'),
      )
      hits += n
      if (edit.commit()) changed++
    }
    return {
      op: op.op,
      matched,
      changed,
      skippedProtected: skipped,
      detail: `${hits} occurrence(s) styled`,
    }
  },
})

function countProtected(entries: SessionEntry[]): number {
  return entries.filter((e) => !TEXT_TYPES.has(entryType(e))).length
}

function validateFontFields(op: Op, where: string): string | null {
  for (const k of ['bold', 'italic', 'underline', 'strike'] as const) {
    if (op[k] !== undefined && typeof op[k] !== 'boolean') return `${where}: ${k} must be a boolean`
  }
  if (op.color !== undefined) {
    const error = checkHex(op.color, where, 'color')
    if (error) return error
  }
  if (op.highlight !== undefined && !isStringOrNull(op.highlight)) {
    return `${where}: highlight must be a color name / hex string or null`
  }
  if (op.fontSize !== undefined && !isPositiveOrNull(op.fontSize)) {
    return `${where}: fontSize must be a positive number of points (or null to clear)`
  }
  if (op.fontFamily !== undefined && !isStringOrNull(op.fontFamily)) {
    return `${where}: fontFamily must be a string or null`
  }
  if (
    op.baseline !== undefined &&
    op.baseline !== null &&
    !(BASELINES as readonly string[]).includes(String(op.baseline))
  ) {
    return `${where}: baseline must be superscript / subscript / none (or null)`
  }
  if (present(op, FONT_KEYS).length === 0) {
    return `${where}: give at least one font field (${FONT_KEYS.join(', ')})`
  }
  return null
}

function validateParagraphFields(op: Op, where: string): string | null {
  if (
    op.align !== undefined &&
    op.align !== null &&
    !(ALIGNS as readonly string[]).includes(String(op.align))
  ) {
    return `${where}: align must be left / center / right / justify (or null)`
  }
  for (const k of [
    'lineSpacing',
    'indentLeft',
    'indentRight',
    'indentFirstLine',
    'spaceBefore',
    'spaceAfter',
  ] as const) {
    if (op[k] !== undefined && !isNumberOrNull(op[k]))
      return `${where}: ${k} must be a number or null`
  }
  if (op.pageBreakBefore !== undefined && typeof op.pageBreakBefore !== 'boolean') {
    return `${where}: pageBreakBefore must be a boolean`
  }
  if (op.shadingFill !== undefined) {
    const error = checkHex(op.shadingFill, where, 'shadingFill')
    if (error) return error
  }
  if (op.borders !== undefined && op.borders !== null) {
    if (typeof op.borders !== 'string' || !/^[tblr]*$/.test(op.borders)) {
      return `${where}: borders must be a subset of "tblr" (or null)`
    }
  }
  if (present(op, PARA_KEYS).length === 0) {
    return `${where}: give at least one paragraph field (${PARA_KEYS.join(', ')})`
  }
  return null
}

register({
  name: 'setParagraphFormat',
  signature:
    'setParagraphFormat <target> align? lineSpacing? indentLeft? indentRight? indentFirstLine? spaceBefore? spaceAfter? pageBreakBefore? shadingFill? borders?',
  keys: PARA_KEYS,
  target: 'required',
  validate: (op, where) => validateParagraphFields(op, where),
  apply: (op, env) => {
    const { text, skipped, matched } = textIndexes(env, op.target as Target)
    let changed = 0
    for (const index of text) {
      const edit = editEntry(env.entries, index)
      const format: ParaFormat = edit.gen.format ?? (edit.gen.format = {})
      const clear = (key: keyof ParaFormat) => delete format[key]
      const setNum = (key: keyof ParaFormat, value: unknown) => {
        if (value === null) clear(key)
        else (format[key] as number) = value as number
      }
      if (op.align !== undefined) {
        if (op.align === null) clear('align')
        else format.align = op.align as ParaFormat['align']
      }
      if (op.lineSpacing !== undefined) setNum('lineSpacing', op.lineSpacing)
      if (op.indentLeft !== undefined) setNum('indentLeft', op.indentLeft)
      if (op.indentRight !== undefined) setNum('indentRight', op.indentRight)
      if (op.indentFirstLine !== undefined) setNum('indentFirstLine', op.indentFirstLine)
      if (op.spaceBefore !== undefined) {
        if (op.spaceBefore === null) clear('spaceBefore')
        else {
          format.spaceBefore = op.spaceBefore as number
          format.spaceBeforeAuto = false
        }
      }
      if (op.spaceAfter !== undefined) {
        if (op.spaceAfter === null) clear('spaceAfter')
        else {
          format.spaceAfter = op.spaceAfter as number
          format.spaceAfterAuto = false
        }
      }
      if (op.pageBreakBefore !== undefined) format.pageBreakBefore = op.pageBreakBefore === true
      if (op.shadingFill !== undefined) {
        const hex = normalizeHex(op.shadingFill, op.op, 'shadingFill')
        if (hex) format.shadingFill = hex
        else clear('shadingFill')
      }
      if (op.borders !== undefined) {
        if (op.borders === null) clear('borders')
        else format.borders = op.borders as string
      }
      if (edit.commit()) changed++
    }
    return { op: op.op, matched, changed, skippedProtected: skipped }
  },
})

register({
  name: 'setHeadingLevel',
  signature: 'setHeadingLevel <target> level (0-6; 0 = plain paragraph)',
  keys: ['level'],
  target: 'required',
  validate: (op, where) => {
    if (!Number.isInteger(op.level) || Number(op.level) < 0 || Number(op.level) > 6) {
      return `${where}: level must be an integer between 0 and 6 (0 = plain paragraph)`
    }
    return null
  },
  apply: (op, env) => {
    const { text, skipped, matched } = textIndexes(env, op.target as Target)
    const level = Number(op.level)
    let changed = 0
    for (const index of text) {
      const edit = editEntry(env.entries, index)
      if (level === 0) {
        edit.gen.type = 'paragraph'
        delete edit.gen.level
        delete edit.gen.outlineOnly
      } else {
        edit.gen.type = 'heading'
        edit.gen.level = level
      }
      if (edit.commit()) changed++
    }
    // BUG-1631: when the document has no style for the level, the serializer
    // keeps the paragraph's current style and carries the level with a direct
    // w:outlineLvl override — the retag is durable, but the paragraph keeps
    // its previous formatting. Disclose that instead of a bare "changed N".
    const warnings =
      level >= 1 && changed > 0 && env.doc.headingStyleIds?.get(level) === undefined
        ? [
            `no style for heading level ${level} in this document; the level is applied as a direct ` +
              'outline-level override and the paragraphs keep their previous formatting',
          ]
        : undefined
    return {
      op: op.op,
      matched,
      changed,
      skippedProtected: skipped,
      ...(warnings ? { warnings } : {}),
    }
  },
})

register({
  name: 'findReplace',
  signature: 'findReplace find replace matchCase? <target?> — small in-place text fixes',
  keys: ['find', 'replace', 'matchCase'],
  target: 'optional',
  validate: (op, where) => {
    if (typeof op.find !== 'string' || op.find === '')
      return `${where}: find must be a non-empty string`
    if (typeof op.replace !== 'string') return `${where}: replace must be a string`
    return null
  },
  apply: (op, env) => {
    const find = op.find as string
    const replace = op.replace as string
    const matchCase = op.matchCase !== false
    const { text, matched } = op.target
      ? textIndexes(env, op.target)
      : { text: allTextIndexes(env.entries), matched: env.entries.length }
    let changed = 0
    let hits = 0
    for (const index of text) {
      const edit = editEntry(env.entries, index)
      for (const run of edit.gen.runs ?? []) {
        if (matchCase) {
          if (!run.text.includes(find)) continue
          hits += run.text.split(find).length - 1
          run.text = run.text.split(find).join(replace)
        } else {
          // the shared fold-safe replace: lowered indices cannot slice the
          // original (İ expands under toLowerCase, BUG-1101)
          const outcome = replaceCaseInsensitive(run.text, find, replace)
          if (outcome.count === 0) continue
          hits += outcome.count
          run.text = outcome.text
        }
      }
      if (edit.commit()) changed++
    }
    return {
      op: op.op,
      matched,
      changed,
      skippedProtected: 0,
      detail: `${hits} replacement(s)`,
    }
  },
})

function allTextIndexes(entries: SessionEntry[]): number[] {
  return entries.map((_, i) => i).filter((i) => TEXT_TYPES.has(entryType(entries[i]!)))
}

register({
  name: 'deleteBlocks',
  signature: 'deleteBlocks <target>',
  keys: [],
  target: 'required',
  validate: () => null,
  apply: (op, env) => {
    const indexes = matchTarget(env.entries, op.target as Target)
    for (const index of [...indexes].sort((a, b) => b - a)) env.entries.splice(index, 1)
    return { op: op.op, matched: indexes.length, changed: indexes.length, skippedProtected: 0 }
  },
})

register({
  name: 'moveBlocks',
  signature: 'moveBlocks blockIndexes[] afterBlockIndex (-1 = document start)',
  keys: ['blockIndexes', 'afterBlockIndex'],
  target: 'none',
  validate: (op, where) => {
    if (
      !Array.isArray(op.blockIndexes) ||
      op.blockIndexes.length === 0 ||
      op.blockIndexes.some((i) => !Number.isInteger(i) || i < 0)
    ) {
      return `${where}: blockIndexes must be a non-empty array of non-negative integers`
    }
    if (!Number.isInteger(op.afterBlockIndex) || Number(op.afterBlockIndex) < -1) {
      return `${where}: afterBlockIndex must be an integer >= -1 (-1 = document start)`
    }
    return null
  },
  apply: (op, env) => {
    const entries = env.entries
    const unique = [...new Set(op.blockIndexes as number[])].sort((a, b) => a - b)
    for (const index of unique) {
      if (index >= entries.length)
        throw new Error(`moveBlocks: blockIndex ${index} is out of range`)
    }
    const after = op.afterBlockIndex as number
    if (after >= entries.length)
      throw new Error(`moveBlocks: afterBlockIndex ${after} is out of range`)
    const movingSet = new Set(unique)
    const moving = unique.map((i) => entries[i]!)
    // non-moving blocks keep their relative order; "after block N" counts the
    // non-moving blocks at or before the original index N as preceding the gap
    const rest: Array<{ at: number; entry: SessionEntry }> = []
    entries.forEach((entry, i) => {
      if (!movingSet.has(i)) rest.push({ at: i, entry })
    })
    const insertAt = after === -1 ? 0 : rest.filter((r) => r.at <= after).length
    const rebuilt = [
      ...rest.slice(0, insertAt).map((r) => r.entry),
      ...moving,
      ...rest.slice(insertAt).map((r) => r.entry),
    ]
    entries.length = 0
    entries.push(...rebuilt)
    return {
      op: op.op,
      matched: unique.length,
      changed: unique.length,
      skippedProtected: 0,
      detail: `moved to after block ${after}`,
    }
  },
})

register({
  name: 'setList',
  signature: 'setList <target> kind? ("bullet" | "number") — convert blocks to list items',
  keys: ['kind'],
  target: 'required',
  validate: (op, where) => {
    if (op.kind !== undefined && op.kind !== 'bullet' && op.kind !== 'number') {
      return `${where}: kind must be "bullet" or "number"`
    }
    return null
  },
  apply: (op, env) => {
    const kind: 'bullet' | 'ordered' = op.kind === 'number' ? 'ordered' : 'bullet'
    const numId = allocateNumId(kind, env.doc, env.numbering)
    const { text, skipped, matched } = textIndexes(env, op.target as Target)
    let changed = 0
    for (const index of text) {
      if (entryType(env.entries[index]!) === 'listItem') continue
      const edit = editEntry(env.entries, index)
      edit.gen.type = 'listItem'
      edit.gen.list = { kind, numId, ilvl: 0 }
      delete edit.gen.level
      delete edit.gen.outlineOnly
      if (edit.commit()) changed++
    }
    return { op: op.op, matched, changed, skippedProtected: skipped }
  },
})

register({
  name: 'clearList',
  signature: 'clearList <target> — convert list items back to plain blocks',
  keys: [],
  target: 'required',
  validate: () => null,
  apply: (op, env) => {
    const { text, matched } = textIndexes(env, op.target as Target)
    let changed = 0
    for (const index of text) {
      if (entryType(env.entries[index]!) !== 'listItem') continue
      const edit = editEntry(env.entries, index)
      edit.gen.type = 'paragraph'
      delete edit.gen.list
      if (edit.commit()) changed++
    }
    return { op: op.op, matched, changed, skippedProtected: 0 }
  },
})

// ---- validation + atomic execution ----

function validateTarget(target: unknown, where: string): string | null {
  if (!target || typeof target !== 'object') return `${where}: missing target`
  const tg = target as Target
  if (
    tg.nodeType !== undefined &&
    !(NODE_TYPES as readonly string[]).includes(normalizeNodeType(String(tg.nodeType)))
  ) {
    return `${where}: unknown nodeType "${String(tg.nodeType)}" (accepted: ${NODE_TYPE_SPELLINGS.join(', ')})`
  }
  if (
    tg.headingLevel !== undefined &&
    (!Number.isInteger(tg.headingLevel) || tg.headingLevel < 1 || tg.headingLevel > 6)
  ) {
    return `${where}: headingLevel must be an integer between 1 and 6`
  }
  if (
    tg.blockIndexes !== undefined &&
    (!Array.isArray(tg.blockIndexes) || tg.blockIndexes.some((i) => !Number.isInteger(i) || i < 0))
  ) {
    return `${where}: blockIndexes must be an array of non-negative integers`
  }
  const hasCondition =
    tg.nodeType !== undefined ||
    tg.headingLevel !== undefined ||
    typeof tg.containsText === 'string' ||
    (Array.isArray(tg.blockIndexes) && tg.blockIndexes.length > 0)
  if (!hasCondition) return `${where}: target requires at least one condition`
  return null
}

function validateShape(op: Op, def: OpDef, where: string): string | null {
  if (def.target === 'none' && op.target !== undefined) {
    return `${where}: does not take a target`
  }
  if (def.target === 'required' || op.target !== undefined) {
    const error = validateTarget(op.target, where)
    if (error) return error
  }
  const unknown = Object.keys(op).filter(
    (k) => k !== 'op' && k !== 'target' && !def.keys.includes(k),
  )
  if (unknown.length > 0) {
    return `${where}: unknown field(s) ${unknown.join(', ')}; allowed: ${def.keys.join(', ')}`
  }
  return def.validate(op, where)
}

/** deep-copy the editable parts of a session (originals are shared, never mutated) */
function cloneEntries(entries: SessionEntry[]): SessionEntry[] {
  return entries.map((entry) =>
    entry.kind === 'edited' ? { ...entry, gen: cloneGen(entry.gen) } : entry,
  )
}

/**
 * Validate the whole batch first, then apply every op sequentially to a working
 * copy of the entries; returns the new list plus the numbering additions only
 * on success (all-or-nothing — a rejected batch mutates nothing).
 */
export function executeOps(
  entries: SessionEntry[],
  ops: Op[],
  doc: NumberingSource,
): ExecuteOutcome & { entries?: SessionEntry[]; numbering?: PendingNumbering } {
  if (!Array.isArray(ops) || ops.length === 0) {
    return { ok: false, results: [], error: 'ops must be a non-empty array' }
  }
  // validate up front: any invalid op rejects the whole batch
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i]
    const where = op && typeof op.op === 'string' ? `op ${i + 1} (${op.op})` : `op ${i + 1}`
    if (!op || typeof op !== 'object' || typeof op.op !== 'string') {
      return { ok: false, results: [], error: `${where}: must be an object with an "op" name` }
    }
    const def = REGISTRY.get(op.op)
    if (!def) {
      return {
        ok: false,
        results: [],
        error: `${where}: unknown op "${op.op}"; known: ${opNames().join(', ')}`,
      }
    }
    const error = validateShape(op, def, where)
    if (error) return { ok: false, results: [], error }
  }
  const working = cloneEntries(entries)
  const numbering = emptyNumbering()
  const results: OpResult[] = []
  try {
    for (const op of ops) {
      const def = REGISTRY.get(op.op)!
      results.push(def.apply(op, { entries: working, doc, numbering }))
    }
  } catch (e) {
    return { ok: false, results, error: e instanceof Error ? e.message : String(e) }
  }
  return { ok: true, results, entries: working, numbering }
}

/** SaveBlock[] for the engine's saveDocx, with rawPPr passthrough applied. */
export function toSaveBlocks(
  entries: SessionEntry[],
): Array<
  | { kind: 'original'; docxIndex: number }
  | { kind: 'generated'; block: GeneratedBlock }
  | { kind: 'xml'; xml: string }
> {
  const out: Array<
    | { kind: 'original'; docxIndex: number }
    | { kind: 'generated'; block: GeneratedBlock }
    | { kind: 'xml'; xml: string }
  > = []
  for (const entry of entries) {
    if (entry.kind === 'original') {
      if (entry.block.hidden || entry.block.docxIndex === null) continue
      out.push({ kind: 'original', docxIndex: entry.block.docxIndex })
    } else if (entry.kind === 'table') {
      out.push({ kind: 'xml', xml: entry.xml })
    } else {
      applyRawPPr(entry.gen, entry.original)
      out.push({ kind: 'generated', block: entry.gen })
    }
  }
  return out
}

/** view of an entry for read/serialization (block-model shape) */
export function entryBlock(entry: SessionEntry): Block | GeneratedBlock {
  if (entry.kind === 'original') return entry.block
  if (entry.kind === 'table') return { type: 'table', table: entry.model } as Block
  return entry.gen
}

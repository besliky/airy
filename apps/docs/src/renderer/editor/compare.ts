/**
 * Review → Compare: paragraph-level diff of two documents (pure, unit-testable).
 *
 * Two result shapes:
 * - `compareParagraphs` feeds the side-by-side differences panel (ComparePanel).
 * - `mergeCompareDocs` builds the Word "legal blackline": a merged copy of the
 *   current document where content missing from the compared file carries `del`
 *   marks (or a block-level `blockRevision` deletion for blocks that cannot
 *   carry run marks) and content coming only from the compared file carries
 *   `ins` marks — accept/reject then work through the regular revisions engine.
 *
 * First-version scope: the diff keys on TEXT content only. Formatting
 * differences (bold/size/style/paragraph-format changes) are not marked;
 * for unchanged text the current document's formatting wins.
 */
import type { Block } from '@airy-office/docx-engine'
import type { PmMark, PmNode } from './convert'

export interface CompareEntry {
  kind: 'same' | 'removed' | 'added' | 'changed'
  /** paragraph text in the current document */
  left?: string
  /** paragraph text in the compared document */
  right?: string
}

/** visible block -> comparable plain text (tables/objects fall back to previews) */
export function blockTexts(blocks: Block[]): string[] {
  return blocks
    .filter((b) => !b.hidden)
    .map((b) => {
      if (b.runs) return b.runs.map((r) => r.text).join('')
      return b.previewText ?? ''
    })
}

/** LCS-based paragraph diff; a removal directly followed by an addition merges into 'changed' */
export function compareParagraphs(left: string[], right: string[]): CompareEntry[] {
  const n = left.length
  const m = right.length
  // lcs[i][j] = LCS length of left[i:], right[j:]
  const lcs: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0))
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i][j] =
        left[i] === right[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1])
    }
  }
  const raw: CompareEntry[] = []
  let i = 0
  let j = 0
  while (i < n && j < m) {
    if (left[i] === right[j]) {
      raw.push({ kind: 'same', left: left[i], right: right[j] })
      i++
      j++
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      raw.push({ kind: 'removed', left: left[i] })
      i++
    } else {
      raw.push({ kind: 'added', right: right[j] })
      j++
    }
  }
  while (i < n) raw.push({ kind: 'removed', left: left[i++] })
  while (j < m) raw.push({ kind: 'added', right: right[j++] })

  const out: CompareEntry[] = []
  for (const entry of raw) {
    const prev = out[out.length - 1]
    if (entry.kind === 'added' && prev?.kind === 'removed' && prev.right === undefined) {
      prev.kind = 'changed'
      prev.right = entry.right
      continue
    }
    out.push(entry)
  }
  return out
}

export interface CompareSummary {
  added: number
  removed: number
  changed: number
}

export function summarize(entries: CompareEntry[]): CompareSummary {
  return {
    added: entries.filter((e) => e.kind === 'added').length,
    removed: entries.filter((e) => e.kind === 'removed').length,
    changed: entries.filter((e) => e.kind === 'changed').length,
  }
}

// ---- Legal blackline: merge the compared document as tracked changes ----

/** top-level node types whose content is plain inline runs (run-level diffable) */
const TEXT_BLOCK_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
/** inline node types that can carry ins/del marks (schema `marks: '_'`) */
const MARKABLE_TYPES = new Set(['text', 'hardBreak'])
/** container node types walked when flattening a block to comparable text */
const CONTAINER_TYPES = new Set([
  'docParagraph',
  'docHeading',
  'docListItem',
  'docTable',
  'docTableHeader',
  'docTableRow',
  'docTableCell',
  'docNestedTable',
  'docCellBoxes',
])
/**
 * Marks an object run (image, equation, note ref, protected block…) in the
 * flattened block text so two blocks that differ only by an object are not
 * reported as identical (the object would then silently stay/depart).
 */
const ATOM = '\uE000'
/**
 * Token-pair budget for the run-level LCS (chars x chars). Above it the
 * paragraph falls back to a whole-block delete+insert swap, keeping the merge
 * bounded on pathologically long paragraphs.
 */
const RUN_DIFF_BUDGET = 4_000_000

export interface CompareStamp {
  author: string
  date: string
}

export interface CompareMergeResult {
  /** merged top-level nodes for the result document */
  content: PmNode[]
  /** difference counts (paragraph granularity, like the diff panel) */
  summary: CompareSummary
}

/** flattened comparable text of one top-level editor node */
export function pmBlockText(node: PmNode): string {
  const parts: string[] = []
  collectBlockText(node, parts)
  return parts.join('')
}

function collectBlockText(node: PmNode, out: string[]): void {
  if (node.type === 'text') {
    out.push(node.text ?? '')
  } else if (node.type === 'hardBreak') {
    out.push(node.attrs?.pageBreak ? '\f' : node.attrs?.colBreak ? '\v' : '\n')
  } else if (CONTAINER_TYPES.has(node.type)) {
    for (const child of node.content ?? []) collectBlockText(child, out)
  } else {
    // images, equations, ruby, note refs, protected blocks…: one marker per object
    out.push(ATOM + node.type)
  }
}

/** comparable texts of the editor's top-level nodes (LCS input, mirrors blockTexts) */
export function pmBlockTexts(nodes: PmNode[]): string[] {
  return nodes.map(pmBlockText)
}

function revisionMark(kind: 'ins' | 'del', stamp: CompareStamp): PmMark {
  return { type: kind, attrs: { author: stamp.author, date: stamp.date, id: null } }
}

/** strip the revision marks that would stack with a compare mark (pending edits count as final text) */
function withoutRevisionMarks(marks: PmMark[] | undefined): PmMark[] {
  return (marks ?? []).filter((m) => m.type !== 'ins' && m.type !== 'del' && m.type !== 'rprChange')
}

function markInlineNode(node: PmNode, mark: PmMark): PmNode {
  if (!MARKABLE_TYPES.has(node.type)) return node
  return { ...node, marks: [...withoutRevisionMarks(node.marks), mark] }
}

function isMarkableInline(node: PmNode): boolean {
  return MARKABLE_TYPES.has(node.type)
}

function withBlockRevision(node: PmNode, kind: 'ins' | 'del', stamp: CompareStamp): PmNode {
  return {
    ...node,
    attrs: {
      ...node.attrs,
      blockRevision: { kind, author: stamp.author, date: stamp.date },
    },
  }
}

/**
 * A block missing from the compared file: text blocks get their runs struck
 * with `del` marks (visible blackline + whole-block removal on Accept);
 * blocks that cannot carry run marks (tables, objects) or empty blocks get a
 * block-level `blockRevision` deletion instead.
 */
function markDeletedBlock(node: PmNode, stamp: CompareStamp): PmNode {
  const del = revisionMark('del', stamp)
  const content = node.content ?? []
  if (TEXT_BLOCK_TYPES.has(node.type) && content.length > 0 && content.every(isMarkableInline)) {
    return { ...node, content: content.map((child) => markInlineNode(child, del)) }
  }
  return withBlockRevision(node, 'del', stamp)
}

/** A block coming only from the compared file: runs underlined with `ins`, else block-level insertion. */
function markInsertedBlock(node: PmNode, stamp: CompareStamp): PmNode {
  const ins = revisionMark('ins', stamp)
  const content = node.content ?? []
  if (TEXT_BLOCK_TYPES.has(node.type) && content.length > 0 && content.every(isMarkableInline)) {
    return { ...node, content: content.map((child) => markInlineNode(child, ins)) }
  }
  return withBlockRevision(node, 'ins', stamp)
}

/**
 * Sanitize content imported from the compared document so it can live in the
 * current one: docxIndex anchors (they would collide with this document's
 * blocks), revision marks/attrs (the compared file counts as final text),
 * bookmark/comment anchors and note refs (their records are not merged).
 */
function sanitizeForeignNode(node: PmNode): PmNode {
  const attrs: Record<string, unknown> = { ...node.attrs }
  attrs.docxIndex = null
  attrs.pPrChange = null
  attrs.blockRevision = null
  attrs.moveRevision = null
  attrs.paraMarkDel = null
  attrs.bookmarks = null
  attrs.hiddenBookmarks = null
  attrs.commentStarts = null
  attrs.commentEnds = null
  attrs.sdtShell = null
  return {
    ...node,
    attrs,
    marks: withoutRevisionMarks(node.marks).filter((m) => m.type !== 'comment'),
    // note references point at the other file's footnote/endnote records
    content: (node.content ?? [])
      .filter((child) => child.type !== 'docNoteRef')
      .map(sanitizeForeignNode),
  }
}

interface DiffToken {
  /** equality key: one char of text, a break char, or an object marker */
  key: string
  /** source inline node (text tokens share one node, cut on emit) */
  node: PmNode
  /** char offset inside the node's text (undefined for single-token nodes) */
  offset?: number
}

function tokenizeInline(content: PmNode[]): DiffToken[] {
  const tokens: DiffToken[] = []
  for (const node of content) {
    if (node.type === 'text' && node.text) {
      for (let k = 0; k < node.text.length; k++) tokens.push({ key: node.text[k], node, offset: k })
    } else if (node.type === 'hardBreak') {
      tokens.push({
        key: node.attrs?.pageBreak ? '\f' : node.attrs?.colBreak ? '\v' : '\n',
        node,
      })
    } else {
      tokens.push({ key: ATOM + node.type, node })
    }
  }
  return tokens
}

type DiffOp = { op: 'same' | 'del' | 'ins'; start: number; end: number }

/**
 * LCS diff of two token key arrays, as merged spans (`same`/`del` index the
 * left tokens, `ins` the right ones). Same dynamic-programming shape as
 * compareParagraphs, over characters/objects instead of paragraphs.
 */
function diffTokenKeys(a: string[], b: string[]): DiffOp[] {
  const n = a.length
  const m = b.length
  const lcs: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1))
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i][j] = a[i] === b[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1])
    }
  }
  const ops: DiffOp[] = []
  const push = (op: DiffOp['op'], start: number, end: number) => {
    const prev = ops[ops.length - 1]
    if (prev && prev.op === op && prev.end === start) prev.end = end
    else ops.push({ op, start, end })
  }
  let i = 0
  let j = 0
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      push('same', i, i + 1)
      i++
      j++
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      push('del', i, i + 1)
      i++
    } else {
      push('ins', j, j + 1)
      j++
    }
  }
  if (i < n) push('del', i, n)
  if (j < m) push('ins', j, m)
  return ops
}

/** re-emit a token span as inline nodes, cutting shared text nodes at span boundaries */
function emitTokenSpan(
  tokens: DiffToken[],
  from: number,
  to: number,
  mark: PmMark | null,
): PmNode[] {
  const out: PmNode[] = []
  let i = from
  while (i < to) {
    const token = tokens[i]
    if (token.offset === undefined) {
      out.push(mark ? markInlineNode(token.node, mark) : token.node)
      i++
      continue
    }
    let j = i
    while (j < to && tokens[j].node === token.node && tokens[j].offset !== undefined) j++
    const marks = mark ? [...withoutRevisionMarks(token.node.marks), mark] : token.node.marks
    out.push({
      type: 'text',
      text: token.node.text!.slice(token.offset, tokens[j - 1].offset! + 1),
      ...(marks && marks.length > 0 ? { marks } : {}),
    })
    i = j
  }
  return out
}

/**
 * Run-level merge of one changed paragraph pair: keep the left paragraph's
 * formatting and anchor, strike what disappeared, underline what arrived.
 * Falls back to a whole-block swap when either side holds unmarkable objects
 * (images, equations…), when the token diff would exceed the budget, or when
 * nothing survives the diff (a paragraph whose entire content is one revision
 * would trip the accept/reject whole-block heuristic and drop the paragraph).
 */
function mergeChangedBlocks(left: PmNode, right: PmNode, stamp: CompareStamp): PmNode[] {
  const swap = () => [markDeletedBlock(left, stamp), markInsertedBlock(right, stamp)]
  const leftContent = left.content ?? []
  const rightContent = right.content ?? []
  if (left.type !== right.type || !TEXT_BLOCK_TYPES.has(left.type)) return swap()
  if (!leftContent.every(isMarkableInline) || !rightContent.every(isMarkableInline)) return swap()
  const a = tokenizeInline(leftContent)
  const b = tokenizeInline(rightContent)
  if ((a.length + 1) * (b.length + 1) > RUN_DIFF_BUDGET) return swap()
  const ops = diffTokenKeys(
    a.map((tk) => tk.key),
    b.map((tk) => tk.key),
  )
  if (!ops.some((op) => op.op === 'same')) return swap()
  const content: PmNode[] = []
  for (const { op, start, end } of ops) {
    if (op === 'same') content.push(...emitTokenSpan(a, start, end, null))
    else if (op === 'del') content.push(...emitTokenSpan(a, start, end, revisionMark('del', stamp)))
    else content.push(...emitTokenSpan(b, start, end, revisionMark('ins', stamp)))
  }
  return [{ ...left, content }]
}

/**
 * Word-style Compare (legal blackline): walk the paragraph LCS of the two
 * documents and rebuild the current document's top-level nodes with the
 * differences recorded as tracked changes. Unchanged blocks pass through
 * untouched (byte-preserving anchors survive); blocks from the compared file
 * are sanitized (foreign anchors/records stripped) before being inserted.
 */
export function mergeCompareDocs(
  left: PmNode[],
  right: PmNode[],
  stamp: CompareStamp,
): CompareMergeResult {
  const foreign = right.map(sanitizeForeignNode)
  const entries = compareParagraphs(pmBlockTexts(left), pmBlockTexts(foreign))
  const content: PmNode[] = []
  let i = 0
  let j = 0
  for (const entry of entries) {
    if (entry.kind === 'same') {
      content.push(left[i++])
      j++
    } else if (entry.kind === 'removed') {
      content.push(markDeletedBlock(left[i++], stamp))
    } else if (entry.kind === 'added') {
      content.push(markInsertedBlock(foreign[j++], stamp))
    } else {
      content.push(...mergeChangedBlocks(left[i++], foreign[j++], stamp))
    }
  }
  return { content, summary: summarize(entries) }
}

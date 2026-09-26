import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import type { Node as PMNode } from '@tiptap/pm/model'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

/**
 * DOM text-node chunking for giant paragraphs (PERF-1700).
 *
 * prosemirror-view renders one PM text node as exactly one DOM text node and
 * rewrites its `nodeValue` wholesale on every change (TextViewDesc.update).
 * For a multi-megabyte single-line paragraph that resets ~4MB of DOM text per
 * keystroke/undo, and Blink then re-lays-out and re-paints the whole ~700k-px
 * paragraph (measured live: 4.4-5.1s per character, 26.5s for undo x5 —
 * PERF-1647 follow-up / audit 2026-09-24). prosemirror-view has no built-in
 * text-node chunking, but its view synchronizer *does* split a text node into
 * several sibling TextViewDescs wherever a local decoration starts or ends
 * inside it (iterDeco in prosemirror-view/src/viewdesc.js).
 *
 * This plugin plants zero-impact `<wbr>` widget decorations inside paragraphs
 * over PARAGRAPH_CHARS, so the giant text node is mirrored in the DOM as many
 * ~CHUNK_CHARS text nodes. A keystroke then rewrites only the piece containing
 * the caret and Blink's incremental layout/paint only reprocesses that piece's
 * lines.
 *
 * Boundary movement policy: boundaries are positions that follow the text
 * (mapped through every changed transaction), so pieces keep byte-identical
 * text and reuse their TextViewDesc DOM no matter where the edit lands. A
 * piece only grows while the caret keeps typing inside it; once a piece
 * passes REANCHOR_MAX the boundaries are recomputed from each giant
 * paragraph's start — pieces after the caret are rewritten once at that
 * point (a rare sub-second hiccup after ~CHUNK_CHARS characters typed at one
 * spot).
 *
 * The decorations are view-only: widgets never enter the document model, so
 * serialization (getMarkdown/getJSON) is unaffected.
 */

/** Chunk size: measured on a live 4MB paragraph (software raster, xvfb) —
 * 16KiB pieces cut a 1-char tail edit from ~370-450ms down to ~60-70ms;
 * 4KiB pieces measured no further gain. */
const CHUNK_CHARS = 16_384

/** Only paragraphs at least this long get chunked; normal documents never
 * see a single widget decoration (the iterDeco fast path stays active).
 * Shared with giantParagraphA11y, which applies the same "giant paragraph"
 * threshold to the accessibility policy. */
export const GIANT_PARAGRAPH_CHARS = 100_000
const PARAGRAPH_CHARS = GIANT_PARAGRAPH_CHARS

/** A piece is re-anchored (boundaries recomputed) once edits made it longer
 * than this — keeps the per-keystroke DOM write and paint damage bounded. */
const REANCHOR_MAX = CHUNK_CHARS * 2

// Widget identity: Decoration.widget builds a fresh WidgetType on every
// rebuild; prosemirror-view compares widget types with
// `toDOM == other.toDOM && compareObjs(spec, other.spec)` (or spec.key), so
// every widget shares this exact function reference and spec object. That
// keeps WidgetViewDescs (and their DOM) stable across transactions.
const wbrDom = (): HTMLElement => document.createElement('wbr')
const WIDGET_SPEC = { key: 'md-giant-text-chunk', side: 1, raw: true } as const

interface ChunkState {
  /** absolute doc positions of chunk boundaries, sorted, deduped */
  positions: readonly number[]
  decos: DecorationSet
}

const emptyState: ChunkState = { positions: [], decos: DecorationSet.empty }

/** [start, end) content ranges of paragraphs big enough to chunk */
function giantParagraphRanges(doc: PMNode): Array<{ from: number; to: number }> {
  const ranges: Array<{ from: number; to: number }> = []
  doc.descendants((node, pos) => {
    if (
      node.type.name === 'paragraph' &&
      node.isTextblock &&
      node.content.size >= PARAGRAPH_CHARS
    ) {
      ranges.push({ from: pos + 1, to: pos + 1 + node.content.size })
    }
  })
  return ranges
}

// PROTOTYPE (PERF-1700b measurement): hide giant paragraphs from the AX tree

function insideRanges(ranges: Array<{ from: number; to: number }>, pos: number): boolean {
  return ranges.some((r) => pos > r.from && pos < r.to)
}

/** initial boundaries for every giant paragraph: CHUNK_CHARS grid from start */
function gridBoundaries(doc: PMNode): number[] {
  const positions: number[] = []
  for (const { from, to } of giantParagraphRanges(doc)) {
    for (let p = from + CHUNK_CHARS; p < to; p += CHUNK_CHARS) positions.push(p)
  }
  return positions
}

function buildDecos(doc: PMNode, positions: readonly number[]): DecorationSet {
  if (positions.length === 0) return DecorationSet.empty
  const decos: Decoration[] = positions.map((p) => Decoration.widget(p, wbrDom, WIDGET_SPEC))
  return DecorationSet.create(doc, decos)
}

/**
 * True when the boundaries still split every giant paragraph into pieces of
 * bounded size (no gap above REANCHOR_MAX from range start through range
 * end, across consecutive boundaries).
 */
function gapsBounded(
  ranges: Array<{ from: number; to: number }>,
  positions: readonly number[],
): boolean {
  let idx = 0
  for (const { from, to } of ranges) {
    let prev = from
    while (idx < positions.length && positions[idx] < to) {
      if (positions[idx] - prev > REANCHOR_MAX) return false
      prev = positions[idx]
      idx++
    }
    if (to - prev > REANCHOR_MAX) return false
  }
  return idx === positions.length
}

function applyTransaction(tr: Transaction, old: ChunkState): ChunkState {
  if (!tr.docChanged) return old
  const ranges = giantParagraphRanges(tr.doc)
  if (ranges.length === 0) return emptyState

  // map boundaries through the change so they follow the text (pieces keep
  // their content and reuse their DOM); drop duplicates and any position
  // that drifted out of a giant paragraph
  const mapped: number[] = []
  for (const p of old.positions) {
    const m = tr.mapping.map(p, 1)
    if (insideRanges(ranges, m) && mapped[mapped.length - 1] !== m) mapped.push(m)
  }
  if (gapsBounded(ranges, mapped)) {
    return { positions: mapped, decos: buildDecos(tr.doc, mapped) }
  }
  const fresh = gridBoundaries(tr.doc)
  return { positions: fresh, decos: buildDecos(tr.doc, fresh) }
}

export const giantTextChunkingKey = new PluginKey<ChunkState>('mdGiantTextChunk')

export const GiantTextChunking = Extension.create({
  name: 'giantTextChunking',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: giantTextChunkingKey,
        state: {
          init: (_, state) => {
            const positions = gridBoundaries(state.doc)
            return { positions, decos: buildDecos(state.doc, positions) }
          },
          apply: applyTransaction,
        },
        props: {
          decorations(state) {
            const st = giantTextChunkingKey.getState(state)
            if (!st) return DecorationSet.empty
            return st.decos
          },
        },
      }),
    ]
  },
})

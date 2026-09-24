import type { Editor } from '@tiptap/core'
import { findInText, type FindOptions, type FindTarget } from '@airy-office/ui'
import { Fragment, Mark, type Node as PMNode } from '@tiptap/pm/model'
import { searchPluginKey, type SearchRange } from './searchHighlight'

/**
 * BUG-1694: on documents with a giant paragraph (1MB, ~43k matches) painting
 * every match as an inline decoration freezes the UI for many seconds and can
 * kill the renderer. The match LIST stays complete and honest (the find panel
 * counter shows the real total, navigation and replace work on all matches) —
 * only the painted decoration set is capped — plenty for on-screen coverage,
 * while keeping the decorated giant-paragraph rebuild as cheap as possible.
 */
export const MAX_PAINTED_HITS = 500

/**
 * Replace All below this many matches is dispatched inline (the find panel
 * tests rely on it being immediate). Above it, the replace is kicked to the
 * next tick with a cancellation window and status events; either way it runs
 * as ONE transaction (single undo step).
 */
export const REPLACE_ALL_SYNC_LIMIT = 1000

/** max matches merged into one replacement slab inside a textblock */
const SLAB_MATCH_LIMIT = 1000

/** a match: doc positions plus the enclosing textblock start and the offset of the match inside the block's concatenated text */
export interface FindMatch extends SearchRange {
  /** start position of the enclosing textblock node */
  block: number
  /** offset of the match inside the block's flattened text (leaves count as one placeholder char) */
  at: number
}

/** what the host is told about highlight capping and bulk-replace progress */
export type FindStatus =
  | { kind: 'overflow'; unpainted: number }
  | { kind: 'replace'; remaining: number; total: number }
  | { kind: 'idle' }

/** collect matches inside textblocks; inline content is flattened so matches spanning marks are found */
export function findMatches(editor: Editor, query: string, opts: FindOptions): FindMatch[] {
  const found: FindMatch[] = []
  if (!query) return found
  editor.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    let text = ''
    // segments replace the old per-character position array: a 1MB textblock
    // would otherwise build a million-entry array on every rescan
    const segs: { posStart: number; len: number }[] = []
    node.forEach((child, offset) => {
      if (child.isText && child.text) {
        segs.push({ posStart: pos + 1 + offset, len: child.text.length })
        text += child.text
      } else {
        segs.push({ posStart: pos + 1 + offset, len: 1 })
        text += '\0' // leaf placeholder (hard break, math) never matches
      }
    })
    if (segs.length === 0) return false
    // matches arrive in ascending order, so a forward pointer walk is enough
    let segIdx = 0
    let segStart = 0
    const posAt = (offset: number): number => {
      while (segIdx < segs.length && segStart + segs[segIdx]!.len <= offset) {
        segStart += segs[segIdx]!.len
        segIdx++
      }
      const seg = segs[segIdx]
      return seg ? seg.posStart + (offset - segStart) : pos + 1
    }
    for (const i of findInText(text, query, opts)) {
      found.push({
        from: posAt(i),
        to: posAt(i + query.length - 1) + 1,
        block: pos,
        at: i,
      })
    }
    return false
  })
  return found
}

/**
 * Pick the matches to paint when the full list exceeds `limit`: the active
 * match always wins, then the matches the caller flagged as visible (those in
 * the scrolled viewport around the active one), then the earliest ones. The
 * result stays sorted by document position.
 */
export function pickHighlightedRanges(
  ranges: SearchRange[],
  activeIndex: number,
  visibleIndexes: Iterable<number> | undefined,
  limit = MAX_PAINTED_HITS,
): SearchRange[] {
  if (ranges.length <= limit) return ranges
  const picked = new Set<number>()
  const add = (i: number) => {
    if (picked.size < limit && i >= 0 && i < ranges.length) picked.add(i)
  }
  add(activeIndex)
  if (visibleIndexes) for (const i of visibleIndexes) add(i)
  for (let i = 0; picked.size < limit && i < ranges.length; i++) add(i)
  return [...picked].sort((a, b) => ranges[a]!.from - ranges[b]!.from).map((i) => ranges[i]!)
}

/**
 * Matches whose DOM coords sit inside the scroll container, discovered by
 * walking outward from the active match (the one the find panel just visited).
 * prosemirror-view has no viewport API, and per-match probing of a 43k-match
 * document would be its own freeze — the outward walk costs a couple of
 * coordsAtPos calls per on-screen match and stops at the container edges.
 */
function visibleAround(
  editor: Editor,
  anchorIndex: number,
  ranges: SearchRange[],
  limit: number,
): number[] | undefined {
  try {
    let el: HTMLElement | null = editor.view.dom as HTMLElement
    while (el && el.scrollHeight <= el.clientHeight + 1) el = el.parentElement
    const container = el ?? (editor.view.dom as HTMLElement)
    const rect = container.getBoundingClientRect()
    const inside = (pos: number): boolean => {
      const c = editor.view.coordsAtPos(pos)
      return (
        !!c &&
        c.bottom >= rect.top &&
        c.top <= rect.bottom &&
        c.left <= rect.right &&
        c.right >= rect.left
      )
    }
    const anchor = ranges[anchorIndex]
    if (!anchor || !inside(anchor.from)) return undefined
    const found: number[] = [anchorIndex]
    for (let i = anchorIndex + 1; i < ranges.length && found.length < limit; i++) {
      if (!inside(ranges[i]!.from)) break
      found.push(i)
    }
    for (let i = anchorIndex - 1; i >= 0 && found.length < limit; i--) {
      if (!inside(ranges[i]!.from)) break
      found.push(i)
    }
    return found
  } catch {
    return undefined
  }
}

interface Slab {
  /** contentFrom + cut* are the doc positions this slab replaces */
  contentFrom: number
  cutFrom: number
  cutTo: number
  fragment: Fragment
  matches: number
}

/**
 * Split one textblock's matches into replacement slabs. Each slab covers a
 * contiguous span of the block's flattened text and is applied as a SINGLE
 * replace step whose fragment is rebuilt once from the original children —
 * this is what keeps a 43k-match replace on a 1MB paragraph O(block) instead
 * of thousands of position-shifting insert steps that drown the renderer in
 * string churn. Marks are preserved: untouched spans keep their original
 * nodes, replacements inherit the marks at each match start.
 */
function buildSlabs(
  doc: PMNode,
  schema: PMNode['type']['schema'],
  blockStart: number,
  matches: FindMatch[],
  replacement: string,
): Slab[] {
  const block = doc.nodeAt(blockStart)
  if (!block) return []
  const contentFrom = blockStart + 1
  const contentLen = block.content.size
  const children: { node: PMNode; start: number; end: number }[] = []
  let off = 0
  block.forEach((child) => {
    const len = child.isText && child.text ? child.text.length : 1
    children.push({ node: child, start: off, end: off + len })
    off += len
  })
  const slabs: Slab[] = []
  for (let i = 0; i < matches.length; i += SLAB_MATCH_LIMIT) {
    const group = matches.slice(i, i + SLAB_MATCH_LIMIT)
    const first = group[0]!
    const last = group[group.length - 1]!
    const cutFrom = Math.min(first.at, contentLen)
    const cutTo = Math.min(last.at + (last.to - last.from), contentLen)
    if (cutTo <= cutFrom) continue
    const pieces: PMNode[] = []
    let pendingText: string[] = []
    let pendingMarks: readonly Mark[] = []
    const flush = () => {
      if (pendingText.length === 0) return
      const text = pendingText.join('')
      pieces.push(
        pendingMarks.length ? schema.text(text, pendingMarks as Mark[]) : schema.text(text),
      )
      pendingText = []
      pendingMarks = []
    }
    const pushText = (text: string, marks: readonly Mark[]) => {
      if (!text) return
      if (pendingText.length > 0 && Mark.sameSet(pendingMarks as Mark[], marks as Mark[])) {
        pendingText.push(text)
        return
      }
      flush()
      pendingText = [text]
      pendingMarks = marks
    }
    let childIdx = 0
    // emit the original children covering flattened-text range [a, b)
    const emitOriginal = (a: number, b: number) => {
      while (childIdx < children.length && children[childIdx]!.end <= a) childIdx++
      for (let c = childIdx; c < children.length && children[c]!.start < b; c++) {
        const ch = children[c]!
        const s = Math.max(a, ch.start) - ch.start
        const e = Math.min(b, ch.end) - ch.start
        if (ch.node.isText && ch.node.text) pushText(ch.node.text.slice(s, e), ch.node.marks)
        else {
          flush()
          pieces.push(ch.node)
        }
      }
    }
    let cursor = cutFrom
    for (const m of group) {
      emitOriginal(cursor, m.at)
      let h = childIdx
      while (h < children.length && children[h]!.end <= m.at) h++
      const host = children[h]
      pushText(replacement, host && host.node.isText ? host.node.marks : [])
      cursor = m.at + (m.to - m.from)
    }
    emitOriginal(cursor, cutTo)
    flush()
    slabs.push({
      contentFrom,
      cutFrom,
      cutTo,
      fragment: Fragment.from(pieces),
      matches: group.length,
    })
  }
  return slabs
}

export function tiptapFindTarget(editor: Editor): FindTarget & {
  onStatus(listener: (status: FindStatus) => void): () => void
} {
  let ranges: FindMatch[] = []
  const statusListeners = new Set<(status: FindStatus) => void>()
  const emitStatus = (status: FindStatus) => {
    for (const l of statusListeners) l(status)
  }
  const replaceState = { active: false, cancelled: false }

  const paint = (activeIndex: number, all: FindMatch[]) => {
    const visible =
      all.length > MAX_PAINTED_HITS
        ? visibleAround(editor, activeIndex, all, MAX_PAINTED_HITS)
        : undefined
    const painted = pickHighlightedRanges(all, activeIndex, visible)
    editor.view.dispatch(
      editor.state.tr
        .setMeta(searchPluginKey, { ranges: painted, activeIndex })
        .setMeta('uiOnly', true),
    )
    emitStatus({ kind: 'overflow', unpainted: Math.max(0, all.length - painted.length) })
  }

  const stopReplace = () => {
    replaceState.active = false
  }

  return {
    get editable() {
      return editor.isEditable
    },
    search(query, opts, activeIndex) {
      ranges = findMatches(editor, query, opts)
      paint(ranges.length === 0 ? 0 : Math.min(activeIndex, ranges.length - 1), ranges)
      return ranges.length
    },
    activate(index) {
      const range = ranges[index]
      if (!range) return
      paint(index, ranges)
      const { node } = editor.view.domAtPos(range.from)
      const el = node instanceof HTMLElement ? node : node.parentElement
      el?.scrollIntoView({ block: 'center' })
    },
    replaceOne(index, replacement) {
      if (replaceState.active) return
      const m = ranges[index]
      if (!m) return
      editor.commands.command(({ tr }) => {
        tr.insertText(replacement, m.from, m.to)
        return true
      })
    },
    replaceAll(replacement) {
      if (replaceState.active || ranges.length === 0) return
      const { doc, schema } = editor.state
      // group matches per textblock and precompute replacement slabs against
      // the current document
      const byBlock = new Map<number, FindMatch[]>()
      for (const m of ranges) {
        const list = byBlock.get(m.block)
        if (list) list.push(m)
        else byBlock.set(m.block, [m])
      }
      const slabs: Slab[] = []
      for (const [blockStart, ms] of byBlock) {
        ms.sort((a, b) => a.at - b.at)
        slabs.push(...buildSlabs(doc, schema, blockStart, ms, replacement))
      }
      if (slabs.length === 0) return
      // apply from the end of the document backwards so positions of the
      // pending slabs stay valid without mapping
      slabs.sort((a, b) => b.contentFrom + b.cutFrom - (a.contentFrom + a.cutFrom))
      const total = ranges.length
      const apply = () => {
        if (replaceState.cancelled || editor.isDestroyed) {
          stopReplace()
          return
        }
        // one transaction with all precomputed slab steps: the steps are cheap
        // (one replace per slab), so the whole replace costs a single view
        // update and stays one undo step. The painted hits ride along in the
        // same meta — clearing them in a separate dispatch would force a second
        // full rebuild of the (possibly giant) textblock DOM, which on the 1MB
        // fixture doubles the freeze.
        const tr = editor.state.tr
        for (const s of slabs)
          tr.replaceWith(s.contentFrom + s.cutFrom, s.contentFrom + s.cutTo, s.fragment)
        tr.setMeta(searchPluginKey, { ranges: [], activeIndex: 0 })
        editor.view.dispatch(tr)
        stopReplace()
        ranges = []
        emitStatus({ kind: 'idle' })
      }
      if (total <= REPLACE_ALL_SYNC_LIMIT) {
        apply()
        return
      }
      if (total <= REPLACE_ALL_SYNC_LIMIT) {
        apply()
        return
      }
      // BUG-1694: on the 1MB single-paragraph fixture the replace view-update
      // explodes when painted search-hit spans cover the paragraph — measured
      // with in-page frame logging: a 1-match replace = 113ms with no long
      // frame, the 43k-match replace under 500 painted spans = ~12s of
      // continuous main-thread block, and the same replace after dropping the
      // paint in its own dispatch drops to the sub-second class. So drop the
      // paint first (its own dispatch), then replace on the clean textblock.
      // The dispatch still runs on the next tick so the click handler returns
      // and clear() can abort.
      replaceState.active = true
      replaceState.cancelled = false
      paint(0, [])
      emitStatus({ kind: 'replace', remaining: total, total })
      setTimeout(apply, 0)
    },
    clear() {
      replaceState.cancelled = true
      stopReplace()
      ranges = []
      paint(0, [])
      emitStatus({ kind: 'idle' })
    },
    onDocChanged(listener) {
      const onUpdate = ({ transaction }: { transaction: { docChanged: boolean } }) => {
        if (transaction.docChanged) listener()
      }
      editor.on('update', onUpdate)
      return () => {
        editor.off('update', onUpdate)
      }
    },
    onStatus(listener) {
      statusListeners.add(listener)
      return () => {
        statusListeners.delete(listener)
      }
    },
  }
}

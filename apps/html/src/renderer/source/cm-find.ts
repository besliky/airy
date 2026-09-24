import { SearchQuery } from '@codemirror/search'
import { html } from '@codemirror/lang-html'
import { Compartment, StateEffect, StateField, type Text } from '@codemirror/state'
import { Decoration, EditorView, type DecorationSet } from '@codemirror/view'
import type { FindOptions } from '@airy-office/ui'

export interface FindRange {
  from: number
  to: number
}

export const setFindHits = StateEffect.define<{ ranges: FindRange[]; active: number }>()

/**
 * Holds the html language so a bulk replace can swap it out for the duration of
 * the run: thousands of scattered changed ranges invalidate the syntax tree
 * doc-wide, and reparsing between the chunk transactions would redo that work
 * on every chunk. Set once in buildExtensions, reconfigured by setSyntaxSuspended.
 */
export const syntaxCompartment = new Compartment()

/** swap the html language out (suspended) or back in; the tree re-parses lazily */
export function setSyntaxSuspended(view: EditorView, suspended: boolean): void {
  view.dispatch({ effects: syntaxCompartment.reconfigure(suspended ? [] : html()) })
}

/** progress of a running bulk replace; `null` while no replace is in flight */
export interface BulkReplaceProgress {
  total: number
  done: number
}

let bulkReplaceState: BulkReplaceProgress | null = null
const bulkReplaceListeners = new Set<(progress: BulkReplaceProgress | null) => void>()

/**
 * Subscribe to bulk-replace progress. The current state is echoed to every new
 * subscriber first, so a UI subscribing mid-run learns the in-flight totals and
 * an idle UI learns `null`.
 */
export function onBulkReplaceProgress(
  listener: (progress: BulkReplaceProgress | null) => void,
): () => void {
  bulkReplaceListeners.add(listener)
  listener(bulkReplaceState)
  return () => bulkReplaceListeners.delete(listener)
}

/** advance the shared progress state; called by the find target's bulk replace */
export function emitBulkReplaceProgress(progress: BulkReplaceProgress | null): void {
  bulkReplaceState = progress
  for (const listener of bulkReplaceListeners) listener(progress)
}

const hit = Decoration.mark({ class: 'search-hit' })
const activeHit = Decoration.mark({ class: 'search-hit search-hit-active' })

/** find hits painted by the app's own panel (CodeMirror's search panel is not used) */
export const findHighlight = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(deco, tr) {
    deco = deco.map(tr.changes)
    for (const e of tr.effects) {
      if (e.is(setFindHits)) {
        deco = Decoration.set(
          e.value.ranges.map((r, i) =>
            (i === e.value.active ? activeHit : hit).range(r.from, r.to),
          ),
          true,
        )
      }
    }
    return deco
  },
  provide: (f) => EditorView.decorations.from(f),
})

export function collectMatches(doc: Text, query: string, opts: FindOptions): FindRange[] {
  if (!query) return []
  const q = new SearchQuery({
    search: query,
    caseSensitive: opts.matchCase,
    wholeWord: opts.wholeWord,
    literal: true,
  })
  const out: FindRange[] = []
  const cursor = q.getCursor(doc)
  for (let m = cursor.next(); !m.done; m = cursor.next()) {
    out.push({ from: m.value.from, to: m.value.to })
  }
  return out
}

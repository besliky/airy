import { EditorView } from '@codemirror/view'
import type { FindTarget } from '@airy-office/ui'
import {
  collectMatches,
  emitBulkReplaceProgress,
  setFindHits,
  setSyntaxSuspended,
  type FindRange,
} from './cm-find'
import { External } from './cm-setup'

/**
 * Find panel adapter over the source editor. Replacements are plain
 * transactions, so they reach the document buffer like typed edits.
 */
export function cmFindTarget(
  view: EditorView,
  onDocChanged: (listener: () => void) => () => void,
): FindTarget {
  let ranges: FindRange[] = []
  let bulkRunning = false
  const paint = (active: number) => view.dispatch({ effects: setFindHits.of({ ranges, active }) })
  /**
   * Apply every match of the last search as one structural operation. The
   * highlights are dropped first — with hits still painted the replace
   * transaction maps every decorated range through every change, which is
   * O(hits × changes) and froze giant replace-alls for minutes — and the html
   * parser is suspended so the doc-wide tree invalidation cannot start
   * reparsing between the app commit and the preview push. Like CodeMirror's
   * own replace-all, all ranges go out in a single transaction: they refer to
   * the same pre-replace document, history records one undo step, and the app
   * commits once, so one debounced map rebuild + preview push covers the run.
   */
  const replaceAllBulk = (replacement: string) => {
    if (bulkRunning || ranges.length === 0) return
    bulkRunning = true
    try {
      // nothing may stay painted across the bulk changes
      view.dispatch({ effects: setFindHits.of({ ranges: [], active: 0 }) })
      const total = ranges.length
      setSyntaxSuspended(view, true)
      try {
        emitBulkReplaceProgress({ total, done: 0 })
        view.dispatch({
          changes: ranges.map((r) => ({ from: r.from, to: r.to, insert: replacement })),
          userEvent: 'input.replace.all',
        })
        emitBulkReplaceProgress({ total, done: total })
      } finally {
        setSyntaxSuspended(view, false)
      }
      ranges = []
    } finally {
      bulkRunning = false
      emitBulkReplaceProgress(null)
    }
  }
  return {
    get editable() {
      return !view.state.readOnly
    },
    search(query, opts, activeIndex) {
      // mid-replace the buffer is a half-applied intermediate state: rescanning
      // would repaint shifted ranges, so answer with the announced count instead
      if (bulkRunning) return ranges.length
      ranges = collectMatches(view.state.doc, query, opts)
      paint(ranges.length === 0 ? 0 : Math.min(activeIndex, ranges.length - 1))
      return ranges.length
    },
    activate(index) {
      const r = ranges[index]
      if (r === undefined) return
      view.dispatch({
        selection: { anchor: r.from, head: r.to },
        effects: [
          setFindHits.of({ ranges, active: index }),
          EditorView.scrollIntoView(r.from, { y: 'center' }),
        ],
        annotations: [External.of(true)],
      })
    },
    replaceOne(index, replacement) {
      const r = ranges[index]
      if (r === undefined) return
      view.dispatch({
        changes: { from: r.from, to: r.to, insert: replacement },
        userEvent: 'input.replace',
      })
    },
    replaceAll(replacement) {
      replaceAllBulk(replacement)
    },
    clear() {
      ranges = []
      paint(0)
    },
    onDocChanged,
  }
}

import { forwardRef, useEffect, useImperativeHandle, useRef } from 'react'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { redo, redoDepth, undo, undoDepth } from '@codemirror/commands'
import { External, buildExtensions, setLineWrap } from './cm-setup'
import { addAiRanges, clearAiRanges } from './cm-highlight'
import { cmFindTarget } from './find-target'
import { ratioToScrollTop, type ScrollerState } from '../preview/scroll-sync'
import type { FindTarget } from '@airy-office/ui'
import type { Patch } from '../document/patch'

export interface SourceEditorHandle {
  /** replace the whole document without touching the undo history (file load) */
  setDoc(text: string): void
  /** replace the whole document as one undoable step (AI rewrite, rollback) */
  replaceDoc(text: string, highlight: boolean): void
  /** apply validated patches as one undoable step; returns the post-edit ranges */
  applyPatches(patches: readonly Patch[], highlight: boolean): Array<[number, number]>
  clearHighlights(): void
  /** select and scroll a source range into view; focus only when the user asked for the editor */
  revealRange(from: number, to: number, focus?: boolean): void
  /** scroll-sync (UX-1704): move the viewport to a 0..1 proportion of the document */
  applyScrollRatio(ratio: number): void
  undo(): boolean
  redo(): boolean
  canUndo(): boolean
  canRedo(): boolean
  focus(): void
  /** UX-1705: insert clipboard text verbatim at the selection (paste as plain text) */
  insertPlainText(text: string): void
  /** adapter for the find/replace panel; null until the editor is mounted */
  findTarget(): FindTarget | null
}

export interface CursorInfo {
  line: number
  col: number
  /** document offset of the cursor head */
  pos: number
}

interface Props {
  initialText: string
  onChange: (text: string) => void
  onCursor: (cursor: CursorInfo) => void
  /** only the editor's own transactions report changes; External-annotated ones are already known to the caller */
  className?: string
  /** word wrap (UX-1704); toggling reconfigures the live editor, default on as before */
  wordWrap?: boolean
  /** scroll-sync: the scroller's own (non-echo) scroll events surface here as metrics */
  onScroll?: (state: ScrollerState) => void
}

export const SourceEditor = forwardRef<SourceEditorHandle, Props>(function SourceEditor(
  { initialText, onChange, onCursor, className, wordWrap = true, onScroll },
  ref,
) {
  const hostRef = useRef<HTMLDivElement>(null)
  const viewRef = useRef<EditorView | null>(null)
  const findTargetRef = useRef<FindTarget | null>(null)
  const docListeners = useRef(new Set<() => void>())
  const onChangeRef = useRef(onChange)
  const onCursorRef = useRef(onCursor)
  const onScrollRef = useRef(onScroll)
  const wrapRef = useRef(wordWrap)
  onChangeRef.current = onChange
  onCursorRef.current = onCursor
  onScrollRef.current = onScroll

  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    const view = new EditorView({
      parent: host,
      state: EditorState.create({
        doc: initialText,
        extensions: buildExtensions(
          (v) => {
            const head = v.state.selection.main.head
            const line = v.state.doc.lineAt(head)
            onCursorRef.current({ line: line.number, col: head - line.from + 1, pos: head })
          },
          { wrap: wrapRef.current },
        ),
      }),
      dispatchTransactions: (trs, v) => {
        v.update(trs)
        if (trs.some((tr) => tr.docChanged && !tr.annotation(External))) {
          onChangeRef.current(v.state.doc.toString())
        }
        if (trs.some((tr) => tr.docChanged)) for (const l of docListeners.current) l()
      },
    })
    viewRef.current = view
    findTargetRef.current = cmFindTarget(view, (listener) => {
      docListeners.current.add(listener)
      return () => docListeners.current.delete(listener)
    })
    const scrollDOM = view.scrollDOM
    const onScrollEvent = () => {
      onScrollRef.current?.({
        scrollTop: scrollDOM.scrollTop,
        scrollHeight: scrollDOM.scrollHeight,
        clientHeight: scrollDOM.clientHeight,
      })
    }
    scrollDOM.addEventListener('scroll', onScrollEvent)
    return () => {
      scrollDOM.removeEventListener('scroll', onScrollEvent)
      view.destroy()
      viewRef.current = null
      findTargetRef.current = null
    }
    // the document is seeded once; later external replacements go through setDoc
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // word-wrap toggle (UX-1704): reconfigure the compartment in place; the mount
  // already built the editor with wrapRef.current, so the first run is a no-op
  useEffect(() => {
    if (wordWrap === wrapRef.current) return
    wrapRef.current = wordWrap
    if (viewRef.current) setLineWrap(viewRef.current, wordWrap)
  }, [wordWrap])

  useImperativeHandle(ref, () => ({
    setDoc(text) {
      const view = viewRef.current
      if (!view || view.state.doc.toString() === text) return
      view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: text },
        annotations: [External.of(true)],
      })
    },
    replaceDoc(text, highlight) {
      const view = viewRef.current
      if (!view) return
      const len = view.state.doc.length
      view.dispatch({
        changes: { from: 0, to: len, insert: text },
        annotations: [External.of(true)],
        effects: highlight
          ? [clearAiRanges.of(null), addAiRanges.of([[0, text.length]])]
          : [clearAiRanges.of(null)],
      })
    },
    applyPatches(patches, highlight) {
      const view = viewRef.current
      if (!view) return []
      const sorted = [...patches].sort((a, b) => a.from - b.from || a.to - b.to)
      const ranges: Array<[number, number]> = []
      let delta = 0
      for (const p of sorted) {
        ranges.push([p.from + delta, p.from + delta + p.text.length])
        delta += p.text.length - (p.to - p.from)
      }
      view.dispatch({
        changes: sorted.map((p) => ({ from: p.from, to: p.to, insert: p.text })),
        annotations: [External.of(true)],
        effects: highlight ? [addAiRanges.of(ranges)] : [],
      })
      return ranges
    },
    clearHighlights() {
      viewRef.current?.dispatch({ effects: [clearAiRanges.of(null)] })
    },
    revealRange(from, to, focus = true) {
      const view = viewRef.current
      if (!view) return
      const len = view.state.doc.length
      const a = Math.max(0, Math.min(from, len))
      const b = Math.max(a, Math.min(to, len))
      view.dispatch({
        selection: { anchor: a, head: b },
        effects: EditorView.scrollIntoView(a, { y: 'center' }),
        annotations: [External.of(true)],
      })
      if (focus) view.focus()
    },
    applyScrollRatio(ratio) {
      const view = viewRef.current
      if (!view) return
      const dom = view.scrollDOM
      const target = ratioToScrollTop(ratio, {
        scrollTop: dom.scrollTop,
        scrollHeight: dom.scrollHeight,
        clientHeight: dom.clientHeight,
      })
      // the host gate drops our echo, but a same-position write would still
      // fire a pointless scroll event — skip it at the source
      if (target === dom.scrollTop) return
      dom.scrollTop = target
    },
    undo: () => (viewRef.current ? undo(viewRef.current) : false),
    redo: () => (viewRef.current ? redo(viewRef.current) : false),
    canUndo: () => (viewRef.current ? undoDepth(viewRef.current.state) > 0 : false),
    canRedo: () => (viewRef.current ? redoDepth(viewRef.current.state) > 0 : false),
    focus: () => viewRef.current?.focus(),
    insertPlainText: (text) => {
      const view = viewRef.current
      if (!view || !text) return
      const { from, to } = view.state.selection.main
      view.dispatch({
        changes: { from, to, insert: text },
        selection: { anchor: from + text.length },
        userEvent: 'input.paste',
      })
      view.focus()
    },
    findTarget: () => findTargetRef.current,
  }))

  return <div ref={hostRef} className={className} />
})

/**
 * Phased document hydration for giant files (PERF-1647). The first segment is
 * mounted synchronously so the first screen is interactive fast; the remaining
 * segments are appended at the document end, one `markdown.parse` per segment
 * (keeping the parser input small — it is quadratic in input size), yielding
 * to the event loop between chunks so typing keeps working during hydration.
 *
 * Hydration transactions are never undoable (`addToHistory: false`) and never
 * mark the document dirty — they re-assemble the file content that was just
 * read from disk, they are not user edits. Callers that serialize the document
 * (save, export, recovery copy) must wait for hydration to finish or they
 * would persist a partially hydrated body.
 */
import type { Editor } from '@tiptap/core'
import type { Fragment } from '@tiptap/pm/model'
import { skipTrailingNodeMeta } from '@tiptap/extensions'

export const HYDRATION_META = 'mdHydration'

export interface HydrationProgress {
  /** segments fully appended so far (including the synchronous first one) */
  done: number
  total: number
}

export interface HydrationOptions {
  /** between-chunk yield; defaults to a macrotask so pending input is served */
  yieldFn?: () => Promise<void>
  onProgress?: (progress: HydrationProgress) => void
}

export const defaultYield = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0))

/**
 * Separator blank lines that lead an interior segment must not reach the
 * parser: mid-document blank runs are block separators, and fed to a fresh
 * `markdown.parse` they would materialize as an extra empty paragraph and
 * break parity with the monolithic parse of the whole body. (Leading blank
 * lines of the FIRST segment are document-leading blanks — the monolithic
 * parse renders those as an empty paragraph too — so segment 0 is used
 * verbatim.)
 */
function stripLeadingBlankLines(segment: string): string {
  return segment.replace(/^(?:[ \t]*\n)+/, '')
}

function stripTrailingNewline(segment: string): string {
  return segment.replace(/\n$/, '')
}

/**
 * Parse one interior segment and build the PM nodes, dropping the segment's
 * trailing edge artifact: a table delimiter or blank run at a segment's edge
 * produces a space token that the parser materializes as an empty paragraph;
 * mid-document that paragraph does not exist in the monolithic parse of the
 * whole body.
 */
function parseSegmentNodes(
  editor: Editor,
  segment: string,
  trimLead: boolean,
  trimTrail: boolean,
): Fragment | null {
  const manager = editor.markdown
  if (!manager) return null
  let text = segment
  if (trimLead) text = stripLeadingBlankLines(text)
  if (trimTrail) text = stripTrailingNewline(text)
  const json = manager.parse(text)
  let content = json.content
  if (!content?.length) return null
  if (trimTrail) {
    while (content.length > 0) {
      const last = content[content.length - 1]
      const hasVisibleContent = (last.content ?? []).some(
        (inline) =>
          // whitespace-only text paragraphs are the segment-edge artifact this
          // trim exists for; any inline atom (raw-HTML chip BUG-1685, image,
          // hard break, math) is real content the monolithic parse keeps
          (inline.text ?? '').trim() !== '' || inline.type !== 'text',
      )
      if (last.type === 'paragraph' && !hasVisibleContent) {
        content = content.slice(0, -1)
      } else {
        break
      }
    }
    if (!content.length) return null
  }
  return editor.schema.nodeFromJSON({ type: 'doc', content }).content
}

function dispatchHydration(editor: Editor, fragment: Fragment, replaceAll: boolean): void {
  const tr = editor.state.tr
  tr.setMeta('addToHistory', false)
  tr.setMeta(HYDRATION_META, true)
  // interior chunks must not trigger the TrailingNode fill paragraph — it
  // would land mid-document once the next chunks are appended
  tr.setMeta(skipTrailingNodeMeta, true)
  if (replaceAll) {
    tr.replaceWith(0, editor.state.doc.content.size, fragment)
  } else {
    tr.insert(editor.state.doc.content.size, fragment)
  }
  editor.view.dispatch(tr)
}

/**
 * Replace the whole document content with the parsed first segment.
 */
export function mountFirstSegment(editor: Editor, segment: string): boolean {
  const fragment = parseSegmentNodes(editor, segment, false, true)
  if (!fragment) return false
  dispatchHydration(editor, fragment, true)
  return true
}

/**
 * Append segments[1..] to the editor at the document end. Resolves true when
 * every segment was appended, false when the editor was destroyed mid-run.
 */
export async function hydrateSegments(
  editor: Editor,
  segments: readonly string[],
  options: HydrationOptions = {},
): Promise<boolean> {
  const yieldFn = options.yieldFn ?? defaultYield
  const total = segments.length
  for (let i = 1; i < total; i++) {
    if (editor.isDestroyed) return false
    await yieldFn()
    if (editor.isDestroyed) return false
    try {
      const trimTrail = i < total - 1
      const fragment = parseSegmentNodes(editor, segments[i], true, trimTrail)
      if (!fragment) {
        options.onProgress?.({ done: i + 1, total })
        continue
      }
      dispatchHydration(editor, fragment, false)
      joinAdjacentListsAtEnd(editor)
    } catch (err) {
      console.error('[markdown] hydration segment failed:', err)
      return false
    }
    options.onProgress?.({ done: i + 1, total })
  }
  ensureTrailingNode(editor)
  return true
}

/**
 * The monolithic parse leaves a TrailingNode fill paragraph at the document
 * end whenever the body does not end with a paragraph; after a fully skipped
 * hydration the finished document must end the same way.
 */
function ensureTrailingNode(editor: Editor): void {
  const doc = editor.state.doc
  const last = doc.childCount > 0 ? doc.child(doc.childCount - 1) : null
  if (!last || last.type.name === 'paragraph') return
  const tr = editor.state.tr
  tr.setMeta('addToHistory', false)
  tr.setMeta(HYDRATION_META, true)
  tr.insert(doc.content.size, editor.schema.nodes.paragraph?.create() ?? last.type.create())
  editor.view.dispatch(tr)
}

/**
 * Two adjacent segments may each end/start with a list of the same kind (the
 * scanner never cuts inside a list, but a segment boundary between two lists
 * of the same type is legal). The full parse renders them as one list — keep
 * the incremental result bit-equal to the monolithic one by joining.
 */
function joinAdjacentListsAtEnd(editor: Editor): void {
  const doc = editor.state.doc
  const count = doc.childCount
  if (count < 2) return
  const prev = doc.child(count - 2)
  const tail = doc.child(count - 1)
  if (prev.type !== tail.type) return
  if (prev.type.name !== 'bulletList' && prev.type.name !== 'orderedList') return
  if (JSON.stringify(prev.attrs) !== JSON.stringify(tail.attrs)) return
  const pos = doc.content.size - tail.nodeSize
  editor.view.dispatch(editor.state.tr.setMeta('addToHistory', false).join(pos, 1))
}

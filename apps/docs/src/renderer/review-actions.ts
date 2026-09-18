/**
 * Review-tab actions: footnotes/endnotes, comments, revisions, ink
 * annotations and compare (document protection lives in ProtectDialog +
 * App.applyProtectDialog). Extracted from App.tsx; the App component passes a
 * ReviewContext built fresh per call so state never goes stale.
 */
import type { Editor } from '@tiptap/core'
import {
  nextNoteId,
  parseDocx,
  readSections,
  type CommentInfo,
  type NoteInfo,
} from '@airy-office/docx-engine'
import type { Dispatch, SetStateAction } from 'react'
import type { DocState } from './doc-state'
import {
  addCommentToSelection,
  addReplyToCommentRange,
  nextCommentId,
  removeCommentFromDoc,
  wordRangeAtCaret,
} from './editor/comments'
import { blockTexts, diffParagraphs, mergeCompareDocs, type CompareEntry } from './editor/compare'
import { blocksToPmDoc } from './editor/convert'
import { pendingCommentPluginKey } from './editor/extensions'
import type { InkAnnotation } from './editor/ink'
import {
  acceptAllRevisions,
  acceptCurrentRevision,
  collectRevisions,
  rejectAllRevisions,
  rejectCurrentRevision,
  TRACK_IGNORE,
} from './editor/revisions'
import { t } from './i18n/locale'

/** open the note text dialog; id set = editing an existing note */
export interface NotePrompt {
  kind: 'footnote' | 'endnote'
  id?: string
}

/**
 * Author stamped on new comments / revision marks: the shell-configured name
 * (Settings → General, possibly changed live) or the localized default when
 * unset/blank. Kept pure for tests.
 */
export function effectiveAuthorName(configured: string, fallback: string): string {
  const trimmed = configured.trim()
  return trimmed ? trimmed : fallback
}

/** The App state the review actions need; built fresh per call. */
export interface ReviewContext {
  editor: Editor | null
  doc: DocState | null
  dirtyRef: { current: boolean }
  setStatus: (status: string) => void
  /** shell-configured author name ('' / undefined = unset → localized default author) */
  authorName?: string
  notePrompt: NotePrompt | null
  setNotePrompt: (value: NotePrompt | null) => void
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  setFootnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setEndnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setNotesDirty: (dirty: boolean) => void
  comments: CommentInfo[]
  setComments: Dispatch<SetStateAction<CommentInfo[]>>
  setCommentsDirty: (dirty: boolean) => void
  setCommentComposing: (composing: boolean) => void
  setShowComments: (show: boolean) => void
  setInkAnnotations: Dispatch<SetStateAction<InkAnnotation[]>>
  setInksDirty: (dirty: boolean) => void
  setCompareResult: (value: { otherName: string; entries: CompareEntry[] } | null) => void
  /** revision display mode (Review → Show for review); compare forces markup on */
  setRevisionDisplay: (mode: 'all' | 'none' | 'original') => void
}

// ---- References: footnotes / endnotes ----

/** dialog submit: create a new note (+ caret marker) or update an existing one */
export function submitNote(ctx: ReviewContext, text: string): void {
  if (!ctx.notePrompt || !ctx.editor) return
  const { kind, id } = ctx.notePrompt
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  const setList = kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes
  if (id !== undefined) {
    setList(list.map((n) => (n.id === id ? { ...n, text } : n)))
  } else {
    const newId = nextNoteId(list)
    setList([...list, { id: newId, text }])
    ctx.editor
      .chain()
      .focus()
      .insertContent({
        type: 'docNoteRef',
        attrs: { kind, id: newId, num: list.length + 1 },
      } as never)
      .run()
  }
  ctx.setNotesDirty(true)
}

/** delete a note, remove its in-text marker, renumber the remaining markers */
export function deleteNote(ctx: ReviewContext, kind: 'footnote' | 'endnote', id: string): void {
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  const next = list.filter((n) => n.id !== id)
  ;(kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes)(next)
  ctx.setNotesDirty(true)
  if (!ctx.editor) return
  const editor = ctx.editor
  const tr = editor.state.tr
  const removals: Array<{ pos: number; size: number }> = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return
    if (String(node.attrs.id) === id) {
      removals.push({ pos, size: node.nodeSize })
    } else {
      const num = next.findIndex((n) => n.id === String(node.attrs.id)) + 1
      if (num > 0 && num !== node.attrs.num) {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, num })
      }
    }
  })
  for (const { pos, size } of removals.reverse()) tr.delete(pos, pos + size)
  if (tr.docChanged) editor.view.dispatch(tr)
}

// ---- Review: comments / revisions / compare / protection ----

/** keeps the picked range highlighted while the composer holds focus */
export function setPendingCommentRange(
  ctx: ReviewContext,
  range: { from: number; to: number } | null,
): void {
  if (!ctx.editor) return
  ctx.editor.view.dispatch(ctx.editor.state.tr.setMeta(pendingCommentPluginKey, range))
}

export function cancelNewComment(ctx: ReviewContext): void {
  ctx.setCommentComposing(false)
  setPendingCommentRange(ctx, null)
}

/** New comment: open the pane with the composer; the mark is applied on submit */
export function startNewComment(ctx: ReviewContext): void {
  const editor = ctx.editor
  if (!editor) return
  if (editor.state.selection.empty) {
    // Word anchors on the word under a collapsed caret rather than refusing
    const word = wordRangeAtCaret(editor)
    if (!word) {
      ctx.setStatus(t('appSelectTextToComment'))
      return
    }
    editor.commands.setTextSelection(word)
  }
  const { from, to } = editor.state.selection
  setPendingCommentRange(ctx, { from, to })
  ctx.setShowComments(true)
  ctx.setCommentComposing(true)
}

export function submitNewComment(ctx: ReviewContext, text: string): void {
  if (!ctx.editor) return
  setPendingCommentRange(ctx, null)
  const id = nextCommentId(ctx.comments)
  if (!addCommentToSelection(ctx.editor, id)) {
    ctx.setStatus(t('appCommentSelectionLost'))
    ctx.setCommentComposing(false)
    return
  }
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [
    ...prev,
    {
      id,
      author: effectiveAuthorName(ctx.authorName ?? '', t('editorDefaultAuthor')),
      date: now,
      text,
    },
  ])
  ctx.setCommentsDirty(true)
  ctx.setCommentComposing(false)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentAdded'))
}

/** Reply to a comment: the new entry carries parentId; the anchor shares the parent comment's range */
export function replyToComment(
  ctx: ReviewContext,
  parentId: string,
  text: string,
  author?: string,
): boolean {
  if (!ctx.editor) return false
  const id = nextCommentId(ctx.comments)
  if (!addReplyToCommentRange(ctx.editor, parentId, id)) {
    ctx.setStatus(t('appCommentAnchorGone'))
    return false
  }
  const who = author ?? effectiveAuthorName(ctx.authorName ?? '', t('editorDefaultAuthor'))
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [...prev, { id, author: who, date: now, text, parentId }])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentReplied'))
  return true
}

/** Word: comment text edits in place; the author, date and anchor stay */
export function editComment(ctx: ReviewContext, id: string, text: string): void {
  ctx.setComments((prev) => prev.map((c) => (c.id === id ? { ...c, text } : c)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentEdited'))
}

/** Resolve/reopen: the whole thread (parent + replies) gets done set together */
export function resolveComment(ctx: ReviewContext, id: string, done: boolean): void {
  ctx.setComments((prev) =>
    prev.map((c) => (c.id === id || c.parentId === id ? { ...c, done } : c)),
  )
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(done ? t('appCommentResolvedMsg') : t('appCommentReopenedMsg'))
}

export function deleteComment(ctx: ReviewContext, id: string): void {
  if (!ctx.editor) return
  // cascade: deleting a parent comment also deletes its replies (including their anchor marks)
  const victims = [id, ...ctx.comments.filter((c) => c.parentId === id).map((c) => c.id)]
  for (const v of victims) removeCommentFromDoc(ctx.editor, v)
  ctx.setComments((prev) => prev.filter((c) => !victims.includes(c.id)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
}

export function handleRevision(
  ctx: ReviewContext,
  action: 'accept' | 'reject',
  all: boolean,
): void {
  if (!ctx.editor) return
  if (all) {
    if (action === 'accept') acceptAllRevisions(ctx.editor)
    else rejectAllRevisions(ctx.editor)
    ctx.setStatus(action === 'accept' ? t('appAllRevisionsAccepted') : t('appAllRevisionsRejected'))
  } else {
    const ok =
      action === 'accept' ? acceptCurrentRevision(ctx.editor) : rejectCurrentRevision(ctx.editor)
    if (!ok) ctx.setStatus(t('appNoRevisionsToHandle'))
  }
  ctx.dirtyRef.current = true
}

// ---- Draw: overlay annotations ----

export function addInk(ctx: ReviewContext, annotation: InkAnnotation): void {
  ctx.setInkAnnotations((prev) => [...prev, annotation])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function removeInks(ctx: ReviewContext, ids: string[]): void {
  const gone = new Set(ids)
  ctx.setInkAnnotations((prev) => prev.filter((a) => !gone.has(a.id)))
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function clearInks(ctx: ReviewContext): void {
  ctx.setInkAnnotations([])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appInksCleared'))
}

/**
 * Compare: pick a second .docx and diff it against the open document.
 * - 'panel' shows the paragraph-level differences pane (unchanged behavior;
 *   read-only safe — it diffs the SAVED document `doc.parsed.blocks`)
 * - 'merge' builds the Word legal blackline: the current document's content is
 *   rebuilt with the differences recorded as tracked changes, ready for the
 *   regular accept/reject machinery (Review tab)
 *
 * Known mode divergence: panel diffs the saved state while merge diffs the
 * live editor (`editor.getJSON()`), so unsaved edits make the two modes
 * disagree until the document is saved (documented, BUG-915 audit note).
 */
export async function compareWithFile(ctx: ReviewContext, mode: 'panel' | 'merge'): Promise<void> {
  if (!ctx.doc) return
  const editor = ctx.editor
  if (mode === 'merge') {
    if (!editor) return
    // UX-903: the merge rebuilds the whole document via a programmatic
    // dispatch, which setEditable(false) alone does not fence — a read-only
    // editor (Restrict Editing / write lock / Read Mode) must refuse here
    if (!editor.isEditable) {
      ctx.setStatus(t('reviewCompareReadonly'))
      return
    }
    // BUG-915: refuse to stack a second blackline over pending revisions —
    // the paragraph keys would count struck/underlined text as plain text,
    // the pairing slides and spans get re-stamped, mixing authors/dates from
    // two sessions (Word instead offers to discard pending changes; an
    // honest refusal is the cheap correct option)
    if (collectRevisions(editor.state.doc).length > 0) {
      ctx.setStatus(t('reviewComparePendingRevisions'))
      return
    }
  }
  const other = await window.desktop.openDocx()
  if (!other) return
  // password-protected comparison target: not wired through the decrypt prompt (yet)
  if ('needsPassword' in other) {
    ctx.setStatus(t('appCompareFailed', { error: t('appDocPwdTitle') }))
    return
  }
  try {
    const otherParsed = await parseDocx(new Uint8Array(other.data))
    if (mode === 'panel') {
      const { entries, degraded } = diffParagraphs(
        blockTexts(ctx.doc.parsed.blocks),
        blockTexts(otherParsed.blocks),
      )
      ctx.setCompareResult({ otherName: other.name, entries })
      // BUG-913: above the paragraph-LCS cell budget the pairing is positional
      if (degraded) ctx.setStatus(t('reviewCompareDegraded'))
      return
    }
    // panel mode returned above; only the merge path continues (editor guarded above)
    if (!editor) return
    const { content, summary, degraded } = mergeCompareDocs(
      editor.getJSON().content ?? [],
      blocksToPmDoc(otherParsed.blocks, readSections(otherParsed)).content ?? [],
      {
        author: effectiveAuthorName(ctx.authorName ?? '', t('editorDefaultAuthor')),
        date: new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
      },
    )
    if (summary.added + summary.removed + summary.changed === 0) {
      ctx.setStatus(t('reviewCompareIdentical', { name: other.name }))
      return
    }
    editor.view.dispatch(
      editor.state.tr
        .replaceWith(
          0,
          editor.state.doc.content.size,
          content.map((node) => editor.schema.nodeFromJSON(node)),
        )
        // the recorder would re-record the whole rebuilt document as one edit
        .setMeta(TRACK_IGNORE, true),
    )
    ctx.setRevisionDisplay('all')
    ctx.dirtyRef.current = true
    ctx.setStatus(
      degraded
        ? t('reviewCompareMergedApprox', {
            name: other.name,
            added: summary.added,
            removed: summary.removed,
            changed: summary.changed,
          })
        : t('reviewCompareMerged', {
            name: other.name,
            added: summary.added,
            removed: summary.removed,
            changed: summary.changed,
          }),
    )
  } catch (err) {
    ctx.setStatus(t('appCompareFailed', { error: String(err) }))
  }
}

/// Threaded-comment model and renderer-side store (modern Excel comments).
///
/// The file model lives in the workbook file (`sheets[].threadedComments`,
/// flat per message) and the save pipeline replaces each sheet's whole thread
/// set (see xlsx-threaded-comments.ts in the gateway). Unlike legacy notes,
/// threaded comments have no Univer model without the proprietary preset, so
/// the live state lives in this store: installed from the opened file, mutated
/// through the commands the panel issues, and snapshotted back into the save
/// request. Sheet removal prunes the store so a save never resurrects threads
/// of a deleted sheet.

import type {
  ThreadedCommentMessage,
  WorkbookFile,
  WorkbookThreadedCommentState,
} from '../shared/desktop-api'
import { shiftLinePosition, toRowColumnShift } from './edit-journal'
import type { StructuralJournalOp } from './edit-journal'

export type { ThreadedCommentMessage }

/// One message as displayed and saved. The open wire adds per-message row and
/// column (resolved sidecar-side); they collapse into the thread anchor here.
export interface ThreadedCommentThread {
  readonly id: string
  readonly row: number
  readonly column: number
  readonly resolved: boolean
  /// Root message first, replies in conversation order.
  readonly messages: readonly ThreadedCommentMessage[]
}

/// Author recorded for comments created in Airy. The app has no account
/// system, so a stable app-local identity is used; Excel shows its display
/// name for this person and keeps it on round-trip.
export const AIRY_COMMENT_PERSON_ID = '{8B3E1C42-52A7-4C0E-9D1A-0F6B2C9D5E11}'
export const AIRY_COMMENT_AUTHOR = 'Airy user'

let threadsBySheet = new Map<string, ThreadedCommentThread[]>()
let fileHadThreads = false
const listeners = new Set<() => void>()

/// Threads whose anchor row/column was inside a removed span (Excel deletes a
/// comment with its cell). Kept so Univer's undo of that removal — the exact
/// inverse insert mutation — can restore them; the stack mirrors Univer's
/// LIFO undo order across sheets. Cleared whenever a workbook (re)installs.
interface AnchorTombstone {
  readonly sheetId: string
  readonly axis: 'row' | 'column'
  /// The removed span in the coordinate space current AFTER the removal: the
  /// index a matching inverse insert names, shifted by every later line op.
  readonly index: number
  readonly count: number
  readonly threads: readonly ThreadedCommentThread[]
}

const anchorTombstones: AnchorTombstone[] = []

function notify(): void {
  for (const listener of listeners) listener()
}

export function subscribeThreadedComments(listener: () => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

/// Snapshot object identity changes only when threads mutate, so React
/// effects keyed on it re-run exactly on comment edits.
export function getThreadedCommentsSnapshot(): ReadonlyMap<
  string,
  readonly ThreadedCommentThread[]
> {
  return threadsBySheet
}

export function threadsForSheet(
  sheetId: string | null | undefined,
): readonly ThreadedCommentThread[] {
  if (!sheetId) return []
  return threadsBySheet.get(sheetId) ?? []
}

/// Live lookup for one thread (used by cell indicators so their rendered
/// state — resolved badge, existence — always mirrors the store).
export function getThread(sheetId: string, threadId: string): ThreadedCommentThread | null {
  return threadsBySheet.get(sheetId)?.find((thread) => thread.id === threadId) ?? null
}

export function fileHadThreadedComments(): boolean {
  return fileHadThreads
}

/// Installs the file's threads (replacing any live state). Must run on every
/// workbook open, before the store can be mutated.
export function installWorkbookThreadedComments(file: WorkbookFile): void {
  fileHadThreads = false
  anchorTombstones.length = 0
  const next = new Map<string, ThreadedCommentThread[]>()
  for (const sheet of file.sheets) {
    if (sheet.threadedComments.length > 0) {
      fileHadThreads = true
      next.set(sheet.id, threadsFromOpenMessages(sheet.threadedComments))
    }
  }
  threadsBySheet = next
  notify()
}

/// One message as delivered on open: the wire carries the anchor cell on every
/// message (roots verbatim from the ref, replies resolved sidecar-side).
type OpenThreadedCommentMessage = WorkbookFile['sheets'][number]['threadedComments'][number]

/// Groups flat open-wire messages into threads: roots anchor a thread, replies
/// attach under their closest anchored ancestor (transitively, with a depth
/// cap that also breaks cyclic parts). Orphan replies are dropped. Document
/// order is preserved for the thread list and each reply chain.
export function threadsFromOpenMessages(
  messages: readonly OpenThreadedCommentMessage[],
): ThreadedCommentThread[] {
  const byId = new Map<string, OpenThreadedCommentMessage>()
  for (const message of messages) byId.set(message.id, message)
  interface Draft {
    readonly id: string
    readonly row: number
    readonly column: number
    readonly resolved: boolean
    messages: ThreadedCommentMessage[]
  }
  const drafts: Draft[] = []
  const draftById = new Map<string, Draft>()
  for (const message of messages) {
    if (message.parentId === undefined) {
      // A root needs its own anchor; unparsable refs never reach the wire.
      const draft: Draft = {
        id: message.id,
        row: message.row,
        column: message.column,
        resolved: message.done,
        messages: [asStoredMessage(message)],
      }
      drafts.push(draft)
      draftById.set(draft.id, draft)
      continue
    }
    let cursor: string | null | undefined = message.parentId
    for (let depth = 0; depth < 64 && cursor !== null && cursor !== undefined; depth += 1) {
      const host = draftById.get(cursor)
      if (host) {
        host.messages.push(asStoredMessage(message))
        break
      }
      cursor = byId.get(cursor)?.parentId
    }
  }
  return drafts.map((draft) => ({
    id: draft.id,
    row: draft.row,
    column: draft.column,
    resolved: draft.resolved,
    messages: draft.messages,
  }))
}

function asStoredMessage(message: OpenThreadedCommentMessage): ThreadedCommentMessage {
  return {
    id: message.id,
    personId: message.personId,
    author: message.author,
    text: message.text,
    createdAt: message.created,
    ...(message.userId === undefined ? {} : { userId: message.userId }),
    ...(message.providerId === undefined ? {} : { providerId: message.providerId }),
  }
}

function nowIso(): string {
  return new Date().toISOString()
}

function newId(): string {
  const uuid =
    typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID()
      : `${Date.now().toString(16)}-4c0e-4a7b-9d1a-${Math.random().toString(16).slice(2, 14)}`
  return `{${uuid.toUpperCase()}}`
}

function personFor(): Pick<ThreadedCommentMessage, 'personId' | 'author'> {
  return { personId: AIRY_COMMENT_PERSON_ID, author: AIRY_COMMENT_AUTHOR }
}

function updateSheet(
  sheetId: string,
  update: (threads: ThreadedCommentThread[]) => ThreadedCommentThread[],
): void {
  const threads = threadsBySheet.get(sheetId) ?? []
  const next = new Map(threadsBySheet)
  next.set(sheetId, update(threads))
  threadsBySheet = next
  notify()
}

export function addThread(
  sheetId: string,
  row: number,
  column: number,
  text: string,
): ThreadedCommentThread {
  const thread: ThreadedCommentThread = {
    id: newId(),
    row,
    column,
    resolved: false,
    messages: [{ id: newId(), text, createdAt: nowIso(), ...personFor() }],
  }
  updateSheet(sheetId, (threads) => insertSorted([...threads, thread]))
  return thread
}

export function appendReply(sheetId: string, threadId: string, text: string): void {
  updateSheet(sheetId, (threads) =>
    threads.map((thread) =>
      thread.id === threadId
        ? {
            ...thread,
            messages: [
              ...thread.messages,
              { id: newId(), text, createdAt: nowIso(), ...personFor() },
            ],
          }
        : thread,
    ),
  )
}

export function setThreadResolved(sheetId: string, threadId: string, resolved: boolean): void {
  updateSheet(sheetId, (threads) =>
    threads.map((thread) => (thread.id === threadId ? { ...thread, resolved } : thread)),
  )
}

export function deleteThread(sheetId: string, threadId: string): void {
  updateSheet(sheetId, (threads) => threads.filter((thread) => thread.id !== threadId))
}

/// Drops threads of removed sheets (called alongside save collection so a
/// removed sheet's threads cannot reappear through a later snapshot). Their
/// undo tombstones go with them — a removed sheet has no undo path back.
export function pruneSheets(keptSheetIds: readonly string[]): void {
  const kept = new Set(keptSheetIds)
  let changed = false
  const next = new Map<string, ThreadedCommentThread[]>()
  for (const [sheetId, threads] of threadsBySheet) {
    if (!kept.has(sheetId)) {
      changed = true
      continue
    }
    if (threads.length > 0) next.set(sheetId, threads)
  }
  for (let at = anchorTombstones.length - 1; at >= 0; at -= 1) {
    const tomb = anchorTombstones[at]
    if (tomb !== undefined && !kept.has(tomb.sheetId)) anchorTombstones.splice(at, 1)
  }
  if (!changed) return
  threadsBySheet = next
  notify()
}

/// Save snapshot for the given (still existing) sheets. Sheets with no
/// threads are omitted unless `includeEmpty` — needed when the file had
/// threads at open and the user deleted them all, so the save removes the
/// now-empty parts.
export function collectThreadedCommentStates(
  sheetIds: readonly string[],
  includeEmpty: boolean,
): WorkbookThreadedCommentState[] {
  const states: WorkbookThreadedCommentState[] = []
  for (const sheetId of sheetIds) {
    const known = threadsBySheet.has(sheetId)
    const threads = threadsBySheet.get(sheetId) ?? []
    if (threads.length === 0 && (!includeEmpty || !known)) continue
    states.push({
      sheetId,
      threads: threads.map((thread) => ({
        id: thread.id,
        row: thread.row,
        column: thread.column,
        resolved: thread.resolved,
        messages: [...thread.messages],
      })),
    })
  }
  return states
}

function insertSorted(threads: ThreadedCommentThread[]): ThreadedCommentThread[] {
  return [...threads].sort((a, b) => a.row - b.row || a.column - b.column)
}

/// Remaps thread anchors through a structural edit so a comment stays
/// attached to its cell, exactly as Excel keeps comments anchored:
/// - an inserted line at or above the anchor shifts it;
/// - a removed line carrying an anchor buries the whole thread (replies ride
///   their root's thread object) — Excel deletes the comment with the cell;
/// - a whole-line move carries the anchor with its line;
/// - Univer's undo of a removal (the exact inverse insert) restores the
///   buried threads, matching Excel's undo.
///
/// Called from the same journal listener that feeds `recordStructuralOp`, so
/// the store tracks the same post-operation coordinate space as the journal's
/// cells and hyperlinks; the save snapshot therefore serializes the shifted
/// anchors. Rectangular replace-moves are out of scope: Excel carries comments
/// with the moved cells there too, but the move gate and its overwrite
/// semantics deserve their own pass.
export function remapThreadedCommentAnchors(sheetId: string, op: StructuralJournalOp): void {
  if (!('index' in op)) return
  const shift = toRowColumnShift(op)
  const axis = shift.axis
  const anchorPosition = (thread: ThreadedCommentThread): number =>
    axis === 'row' ? thread.row : thread.column
  const sameTomb = (tomb: AnchorTombstone): boolean =>
    tomb.sheetId === sheetId && tomb.axis === axis

  // Undo: an insert that exactly reverses the newest tombstone pops it and
  // returns its threads to the recreated lines. Buried anchors are stored in
  // the pre-removal space, and the inverse insert restores exactly that space
  // (Univer's undo is strictly LIFO, so intervening ops are undone first),
  // while the surviving anchors shift with the insert.
  const top = anchorTombstones[anchorTombstones.length - 1]
  if (
    top !== undefined &&
    !shift.removing &&
    shift.swap === undefined &&
    sameTomb(top) &&
    top.index === op.index &&
    top.count === op.count
  ) {
    anchorTombstones.pop()
    const resurrected = top.threads
    updateSheet(sheetId, (threads) =>
      insertSorted([...threads.map((thread) => shiftedThread(thread, shift)), ...resurrected]),
    )
    return
  }

  const buried: ThreadedCommentThread[] = []
  const sheetThreads = threadsBySheet.get(sheetId)
  if (sheetThreads !== undefined && sheetThreads.length > 0) {
    const next: ThreadedCommentThread[] = []
    let changed = false
    for (const thread of sheetThreads) {
      const from = anchorPosition(thread)
      const moved = shiftLinePosition(from, shift)
      if (moved === null) {
        buried.push(thread)
        changed = true
      } else if (moved === from) {
        next.push(thread)
      } else {
        next.push(shiftedAnchor(thread, axis, moved))
        changed = true
      }
    }
    if (changed) updateSheet(sheetId, () => insertSorted(next))
  }

  // Tombstones on the same sheet+axis track the moving coordinate space so a
  // later inverse insert still names the right index. One whose recreation
  // span a second removal cuts into is dropped: nothing consistent is left to
  // resurrect into.
  for (let at = anchorTombstones.length - 1; at >= 0; at -= 1) {
    const tomb = anchorTombstones[at]
    if (tomb === undefined || !sameTomb(tomb)) continue
    if (shift.removing) {
      const intersects =
        tomb.index < shift.index + shift.count && tomb.index + tomb.count > shift.index
      if (intersects) {
        anchorTombstones.splice(at, 1)
        continue
      }
    }
    const moved = shiftLinePosition(tomb.index, shift)
    if (moved !== null && moved !== tomb.index) {
      anchorTombstones[at] = { ...tomb, index: moved }
    }
  }
  if (buried.length > 0) {
    anchorTombstones.push({
      sheetId,
      axis,
      index: op.index,
      count: op.count,
      threads: buried,
    })
  }
}

function shiftedThread(
  thread: ThreadedCommentThread,
  shift: ReturnType<typeof toRowColumnShift>,
): ThreadedCommentThread {
  const axis = shift.axis
  const from = axis === 'row' ? thread.row : thread.column
  const moved = shiftLinePosition(from, shift)
  if (moved === null || moved === from) return thread
  return shiftedAnchor(thread, axis, moved)
}

function shiftedAnchor(
  thread: ThreadedCommentThread,
  axis: 'row' | 'column',
  moved: number,
): ThreadedCommentThread {
  return axis === 'row' ? { ...thread, row: moved } : { ...thread, column: moved }
}

import { beforeEach, describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import { buildThreadedCommentsXml } from '../src/gateway/xlsx-threaded-comments'
import type { StructuralJournalOp } from '../src/renderer/edit-journal'
import {
  addThread,
  appendReply,
  collectThreadedCommentStates,
  installWorkbookThreadedComments,
  pruneSheets,
  remapThreadedCommentAnchors,
  subscribeThreadedComments,
  threadsForSheet,
} from '../src/renderer/threaded-comments'

/// BUG-1714 (SC-04): structural row/column edits must remap threaded-comment
/// anchors exactly like Excel keeps comments attached to their cells — an
/// insert above shifts the anchor, deleting the anchor row buries the whole
/// thread, and Univer's undo of the removal brings it back. The save snapshot
/// serializes whatever the store holds, so shifted anchors must reach the
/// threadedComments part as the new refs.
let notifications = 0
subscribeThreadedComments(() => {
  notifications += 1
})

describe('threaded comment anchor remap', () => {
  beforeEach(() => {
    notifications = 0
    // Reset the module store through its public install path.
    installWorkbookThreadedComments({ sheets: [] } as unknown as WorkbookFile)
  })

  const seedSheet = (): { b5: string; b10: string } => {
    // B5 = row 4 / column 1, B10 = row 9 / column 1; one reply rides B5.
    const b5 = addThread('s1', 4, 1, 'on B5').id
    const b10 = addThread('s1', 9, 1, 'on B10').id
    appendReply('s1', b5, 'reply on B5 thread')
    notifications = 0
    return { b5, b10 }
  }

  it('shifts anchors when rows are inserted above and keeps replies on their root', () => {
    const { b5, b10 } = seedSheet()
    // Insert one row at 3 (row index 2): B5 → B6, B10 → B11.
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 2, count: 1 })
    expect(threadsForSheet('s1').map((thread) => [thread.row, thread.column])).toEqual([
      [5, 1],
      [10, 1],
    ])
    // Replies stay inside their thread; only the anchor moved.
    const moved = threadsForSheet('s1').find((thread) => thread.id === b5)!
    expect(moved.messages.map((message) => message.text)).toEqual(['on B5', 'reply on B5 thread'])
    expect(threadsForSheet('s1').find((thread) => thread.id === b10)!.row).toBe(10)
    expect(notifications).toBeGreaterThan(0)
  })

  it('shifts columns on column inserts without touching rows', () => {
    seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'insert-cols', index: 0, count: 2 })
    expect(threadsForSheet('s1').map((thread) => [thread.row, thread.column])).toEqual([
      [4, 3],
      [9, 3],
    ])
  })

  it('leaves anchors above the insert point alone', () => {
    seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 5, count: 3 })
    expect(threadsForSheet('s1').map((thread) => thread.row)).toEqual([4, 12])
  })

  it('serializes the shifted anchor into the threadedComments part', () => {
    seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 2, count: 1 })
    const [state] = collectThreadedCommentStates(['s1'], false)
    const xml = buildThreadedCommentsXml(state!.threads)
    expect(xml).toContain('ref="B6"')
    expect(xml).toContain('ref="B11"')
    expect(xml).not.toContain('ref="B5"')
    expect(xml).not.toContain('ref="B10"')
  })

  it('buries the thread whose anchor row is removed, survivors shift up', () => {
    const { b5 } = seedSheet()
    // Delete row 6 (index 5): B6 anchor would be gone — delete B5's row instead.
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    const threads = threadsForSheet('s1')
    expect(threads.find((thread) => thread.id === b5)).toBeUndefined()
    expect(threads.map((thread) => thread.row)).toEqual([8])
    // The save must not resurrect the buried thread: it is out of the snapshot.
    expect(collectThreadedCommentStates(['s1'], true)).toEqual([
      { sheetId: 's1', threads: threads.map((thread) => ({ ...thread })) },
    ])
  })

  it('undo of the removal resurrects the buried thread and restores survivors', () => {
    const { b5 } = seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    // Univer's undo of remove-rows fires the exact inverse insert.
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 4, count: 1 })
    const threads = threadsForSheet('s1')
    expect(threads.map((thread) => [thread.id, thread.row])).toEqual([
      [b5, 4],
      [threads[1]!.id, 9],
    ])
    expect(threads[0]!.messages.map((message) => message.text)).toEqual([
      'on B5',
      'reply on B5 thread',
    ])
  })

  it('redo of the removal buries the thread again', () => {
    const { b5 } = seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 4, count: 1 })
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    expect(threadsForSheet('s1').map((thread) => thread.row)).toEqual([8])
    expect(threadsForSheet('s1').find((thread) => thread.id === b5)).toBeUndefined()
  })

  it('undo of an insert shifts anchors back without resurrecting anything', () => {
    seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 2, count: 1 })
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 2, count: 1 })
    expect(threadsForSheet('s1').map((thread) => thread.row)).toEqual([4, 9])
  })

  it('does not resurrect on an unrelated insert', () => {
    const { b5 } = seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    // A fresh insert elsewhere (not the inverse) must stay a plain shift.
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 0, count: 1 })
    expect(threadsForSheet('s1').find((thread) => thread.id === b5)).toBeUndefined()
    expect(threadsForSheet('s1').map((thread) => thread.row)).toEqual([9])
  })

  it('keeps the tombstone restorable across intervening line ops before undo', () => {
    const { b5 } = seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    // Intervening insert then its own undo (LIFO like Univer's stack): the
    // buried thread must still come back at its original anchor.
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 0, count: 2 })
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 0, count: 2 })
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 4, count: 1 })
    expect(threadsForSheet('s1').map((thread) => thread.row)).toEqual([4, 9])
    expect(threadsForSheet('s1').find((thread) => thread.id === b5)).toBeDefined()
  })

  it('drops the tombstone when a second removal cuts into the buried span', () => {
    const { b5 } = seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 2 })
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 4, count: 1 })
    expect(threadsForSheet('s1').find((thread) => thread.id === b5)).toBeUndefined()
  })

  it('carries anchors with whole-line moves', () => {
    seedSheet()
    // Move rows [0,2) before row 6 (Excel cut + insert cut cells).
    remapThreadedCommentAnchors('s1', { kind: 'move-rows', index: 0, count: 2, before: 6 })
    // B5 (row 4) sits in the displaced lower block: rows 2-5 shift up to 0-3.
    // B10 (row 9) is below the landing spot: unchanged.
    expect(threadsForSheet('s1').map((thread) => [thread.row, thread.column])).toEqual([
      [2, 1],
      [9, 1],
    ])
  })

  it('keeps the store sorted by cell after a remap', () => {
    addThread('s1', 4, 1, 'B5')
    addThread('s1', 2, 0, 'A3')
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 0, count: 5 })
    expect(threadsForSheet('s1').map((thread) => [thread.row, thread.column])).toEqual([
      [7, 0],
      [9, 1],
    ])
  })

  it('remaps only the edited sheet', () => {
    seedSheet()
    addThread('s2', 3, 0, 'other sheet')
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 0, count: 1 })
    expect(threadsForSheet('s2')[0]!.row).toBe(3)
  })

  it('ignores non-line structural ops', () => {
    seedSheet()
    const op: StructuralJournalOp = {
      kind: 'merge-cells',
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    }
    remapThreadedCommentAnchors('s1', op)
    expect(threadsForSheet('s1').map((thread) => thread.row)).toEqual([4, 9])
  })

  it('drops tombstones of pruned sheets and resets on a new install', () => {
    const { b5 } = seedSheet()
    remapThreadedCommentAnchors('s1', { kind: 'remove-rows', index: 4, count: 1 })
    // A new workbook install clears the undo tombstones: stale resurrection
    // must not leak across sessions.
    installWorkbookThreadedComments({ sheets: [] } as unknown as WorkbookFile)
    addThread('s1', 0, 0, 'fresh')
    remapThreadedCommentAnchors('s1', { kind: 'insert-rows', index: 4, count: 1 })
    expect(threadsForSheet('s1').map((thread) => [thread.row, thread.column])).toEqual([[0, 0]])
    expect(b5).toBeDefined()
    // A tombstone of a removed sheet dies with the sheet.
    addThread('gone', 2, 0, 'doomed')
    remapThreadedCommentAnchors('gone', { kind: 'remove-rows', index: 2, count: 1 })
    pruneSheets(['s1'])
    remapThreadedCommentAnchors('gone', { kind: 'insert-rows', index: 2, count: 1 })
    expect(threadsForSheet('gone')).toHaveLength(0)
  })
})

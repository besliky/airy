import { beforeEach, describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import {
  addThread,
  appendReply,
  collectThreadedCommentStates,
  deleteThread,
  fileHadThreadedComments,
  installWorkbookThreadedComments,
  pruneSheets,
  setThreadResolved,
  subscribeThreadedComments,
  threadsForSheet,
  threadsFromOpenMessages,
} from '../src/renderer/threaded-comments'

type OpenMessage = Parameters<typeof threadsFromOpenMessages>[0][number]

function message(overrides: Partial<OpenMessage> & { id: string }): OpenMessage {
  return {
    row: 0,
    column: 0,
    personId: '{p}',
    author: 'Author',
    text: 'text',
    created: '2026-09-23T00:00:00.000Z',
    done: false,
    ...overrides,
  }
}

function fileWith(sheets: { id: string; threadedComments: OpenMessage[] }[]): WorkbookFile {
  return {
    sheets: sheets.map((sheet) => ({ ...sheet, comments: [] })),
  } as unknown as WorkbookFile
}

let notifications = 0

beforeEach(() => {
  notifications = 0
  // Reset the module store through its public install path.
  installWorkbookThreadedComments({ sheets: [] } as unknown as WorkbookFile)
})

describe('threadsFromOpenMessages', () => {
  it('groups replies under their anchored root and keeps document order', () => {
    const threads = threadsFromOpenMessages([
      message({ id: 'root-1', row: 1, column: 2, done: true }),
      message({ id: 'reply-1', parentId: 'root-1', row: 1, column: 2 }),
      message({ id: 'root-2', row: 4, column: 0 }),
    ])
    expect(threads).toHaveLength(2)
    expect(threads[0]).toMatchObject({ id: 'root-1', row: 1, column: 2, resolved: true })
    expect(threads[0]!.messages.map((entry) => entry.id)).toEqual(['root-1', 'reply-1'])
    // The wire's per-message anchor fields collapse into the thread.
    expect(threads[0]!.messages[1]).not.toHaveProperty('row')
    expect(threads[1]).toMatchObject({ id: 'root-2', resolved: false })
  })

  it('follows replies-to-replies transitively', () => {
    const threads = threadsFromOpenMessages([
      message({ id: 'root', row: 0, column: 5 }),
      message({ id: 'reply', parentId: 'root', row: 0, column: 5 }),
      message({ id: 'nested', parentId: 'reply', row: 0, column: 5 }),
    ])
    expect(threads).toHaveLength(1)
    expect(threads[0]!.messages.map((entry) => entry.id)).toEqual(['root', 'reply', 'nested'])
  })

  it('drops orphan replies whose parent never anchored', () => {
    const threads = threadsFromOpenMessages([
      message({ id: 'orphan', parentId: 'missing' }),
      message({ id: 'chain', parentId: 'orphan' }),
    ])
    expect(threads).toHaveLength(0)
  })

  it('maps the wire created field onto createdAt', () => {
    const [thread] = threadsFromOpenMessages([
      message({ id: 'root', created: '2020-01-02T03:04:05.000Z' }),
    ])
    expect(thread!.messages[0]!.createdAt).toBe('2020-01-02T03:04:05.000Z')
  })
})

describe('thread store', () => {
  it('installs from an opened file and notifies subscribers', () => {
    const unsubscribe = subscribeThreadedComments(() => {
      notifications += 1
    })
    installWorkbookThreadedComments(
      fileWith([
        {
          id: 'sheet-1',
          threadedComments: [message({ id: 'root', row: 3, column: 4, done: true })],
        },
      ]),
    )
    unsubscribe()
    expect(fileHadThreadedComments()).toBe(true)
    expect(threadsForSheet('sheet-1')).toHaveLength(1)
    expect(threadsForSheet('sheet-1')[0]!.resolved).toBe(true)
    expect(notifications).toBe(1)
  })

  it('adds, replies, resolves, and deletes threads', () => {
    const added = addThread('s1', 2, 3, 'first')
    appendReply('s1', added.id, 'second reply')
    setThreadResolved('s1', added.id, true)
    const threads = threadsForSheet('s1')
    expect(threads).toHaveLength(1)
    expect(threads[0]!.messages.map((entry) => entry.text)).toEqual(['first', 'second reply'])
    expect(threads[0]!.resolved).toBe(true)
    deleteThread('s1', added.id)
    expect(threadsForSheet('s1')).toHaveLength(0)
  })

  it('keeps threads sorted by cell', () => {
    addThread('s2', 5, 0, 'later')
    addThread('s2', 1, 1, 'earlier')
    expect(threadsForSheet('s2').map((thread) => [thread.row, thread.column])).toEqual([
      [1, 1],
      [5, 0],
    ])
  })

  it('collects save states only for covered sheets unless empty lists are needed', () => {
    addThread('s1', 0, 0, 'kept')
    // File never had threads: uncovered sheets stay uncovered.
    expect(collectThreadedCommentStates(['s1', 's2'], false)).toEqual([
      {
        sheetId: 's1',
        threads: [
          {
            id: threadsForSheet('s1')[0]!.id,
            row: 0,
            column: 0,
            resolved: false,
            messages: threadsForSheet('s1')[0]!.messages,
          },
        ],
      },
    ])
    // File had threads at open: every covered sheet is sent, empties included,
    // so the save removes the stale parts.
    installWorkbookThreadedComments(
      fileWith([{ id: 's9', threadedComments: [message({ id: 'x' })] }]),
    )
    expect(fileHadThreadedComments()).toBe(true)
    deleteThread('s9', 'x')
    const states = collectThreadedCommentStates(['s9', 'missing'], true)
    expect(states).toEqual([{ sheetId: 's9', threads: [] }])
    // Without includeEmpty the emptied sheet is omitted again.
    expect(collectThreadedCommentStates(['s9'], false)).toEqual([])
  })

  it('prunes threads of removed sheets', () => {
    addThread('keep', 0, 0, 'stays')
    addThread('gone', 0, 0, 'dies')
    pruneSheets(['keep'])
    expect(threadsForSheet('keep')).toHaveLength(1)
    expect(threadsForSheet('gone')).toHaveLength(0)
  })

  it('a workbook without threads clears the had-threads flag', () => {
    installWorkbookThreadedComments(
      fileWith([{ id: 's9', threadedComments: [message({ id: 'x' })] }]),
    )
    expect(fileHadThreadedComments()).toBe(true)
    installWorkbookThreadedComments(fileWith([{ id: 's9', threadedComments: [] }]))
    expect(fileHadThreadedComments()).toBe(false)
  })
})

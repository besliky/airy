/** Comments pane pure logic (components/CommentsPane.tsx): grouping the flat
 *  comment list into threads (parents in order, replies nested under their
 *  parent by authorId+idx). The rendered pane is main-process-backed chrome —
 *  only the grouping is pure. */
import { describe, expect, it } from 'vitest'
import type { SlideComment } from '@airy-office/pptx-engine'

import { commentThreads } from '../src/renderer/components/CommentsPane'

const comment = (over: Partial<SlideComment> & { idx: number }): SlideComment => ({
  authorId: 0,
  author: 'A',
  initials: 'A',
  dt: '2026-01-01T00:00:00Z',
  text: '',
  ...over,
})

describe('commentThreads', () => {
  it('nests replies under their parent and keeps parents in list order', () => {
    const head1 = comment({ authorId: 1, idx: 0, text: 'first' })
    const head2 = comment({ authorId: 2, idx: 0, text: 'second' })
    const reply1 = comment({
      authorId: 3,
      idx: 1,
      parentId: { authorId: 1, idx: 0 },
      text: 'r1',
    })
    const reply2 = comment({
      authorId: 4,
      idx: 2,
      parentId: { authorId: 1, idx: 0 },
      text: 'r2',
    })
    // replies listed BEFORE their parent still nest
    const threads = commentThreads([reply1, head1, reply2, head2])
    expect(threads).toEqual([
      { head: head1, replies: [reply1, reply2] },
      { head: head2, replies: [] },
    ])
  })

  it('matches parents by the authorId+idx pair, not just idx', () => {
    const a = comment({ authorId: 1, idx: 0 })
    const b = comment({ authorId: 2, idx: 0 })
    const replyToB = comment({ authorId: 3, idx: 5, parentId: { authorId: 2, idx: 0 } })
    const threads = commentThreads([a, b, replyToB])
    expect(threads.map((t) => t.replies.length)).toEqual([0, 1])
    expect(threads[1]!.replies[0]).toBe(replyToB)
  })

  it('handles an empty list', () => {
    expect(commentThreads([])).toEqual([])
  })
})

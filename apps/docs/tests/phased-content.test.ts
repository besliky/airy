import { describe, expect, it } from 'vitest'
import type { PmNode } from '../src/renderer/editor/convert'
import {
  PHASE1_BLOCKS,
  PHASE_CHUNK_BLOCKS,
  PHASE_CHUNK_MAX_BLOCKS,
  PHASED_MIN_BLOCKS,
  cancelPhasedContent,
  isPhasedContentPending,
  setContentPhased,
  waitForFullContent,
  type PhasedContentHost,
} from '../src/renderer/phased-content'

const docOf = (blocks: number): PmNode => ({
  type: 'doc',
  content: Array.from({ length: blocks }, (_, i) => ({
    type: 'docParagraph',
    attrs: { docxIndex: i },
  })),
})

/** host that records events; appends mark dirty like the editor's onUpdate does */
function makeHost(destroyed = false) {
  const state = {
    mounted: [] as PmNode[],
    events: [] as string[],
    dirty: false,
    loading: false,
  }
  const host: PhasedContentHost = {
    setContent: (d) => {
      state.mounted = [...(d.content ?? [])]
      state.events.push(`set:${state.mounted.length}`)
    },
    appendNodes: (nodes) => {
      state.mounted.push(...nodes)
      state.dirty = true
      state.events.push(`append:${nodes.length}`)
    },
    isDestroyed: () => destroyed,
    resetHistory: () => state.events.push('resetHistory'),
    setLoading: (v) => {
      state.loading = v
      state.events.push(`loading:${v}`)
    },
    getDirty: () => state.dirty,
    setDirty: (v) => {
      state.dirty = v
    },
  }
  return { host, state }
}

/** scheduler the test drains manually, one chunk per drain() call */
function makeScheduler() {
  const queue: Array<() => void> = []
  return {
    schedule: (cb: () => void) => queue.push(cb),
    drain: () => queue.shift()?.(),
    pending: () => queue.length,
  }
}

describe('setContentPhased', () => {
  it('mounts small documents in one pass without the loading flag', () => {
    const { host, state } = makeHost()
    const s = makeScheduler()
    setContentPhased(host, docOf(PHASED_MIN_BLOCKS), s.schedule, s.schedule)
    expect(state.events).toEqual([`set:${PHASED_MIN_BLOCKS}`])
    expect(s.pending()).toBe(0)
  })

  it('streams a large document in chunks and lands the full content', async () => {
    const blocks = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 2 + 7
    const { host, state } = makeHost()
    const s = makeScheduler()
    setContentPhased(host, docOf(blocks), s.schedule, s.schedule)
    expect(state.mounted.length).toBe(PHASE1_BLOCKS)
    expect(state.loading).toBe(true)
    while (s.pending() > 0) s.drain()
    expect(state.mounted.length).toBe(blocks)
    // every block landed exactly once, in order
    expect(state.mounted.map((n) => n.attrs?.docxIndex)).toEqual(
      Array.from({ length: blocks }, (_, i) => i),
    )
    expect(state.loading).toBe(false)
    expect(state.events).toContain('resetHistory')
    await expect(waitForFullContent()).resolves.toBeUndefined()
  })

  it('streaming is not an edit: the dirty flag survives every chunk', () => {
    const { host, state } = makeHost()
    const s = makeScheduler()
    setContentPhased(host, docOf(PHASED_MIN_BLOCKS + 50), s.schedule, s.schedule)
    while (s.pending() > 0) s.drain()
    expect(state.dirty).toBe(false)
  })

  it('cancelPhasedContent drops the pending tail and releases the save gate', async () => {
    const blocks = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 3
    const { host, state } = makeHost()
    const s = makeScheduler()
    setContentPhased(host, docOf(blocks), s.schedule, s.schedule)
    s.drain() // first chunk lands
    const landed = state.mounted.length
    cancelPhasedContent()
    while (s.pending() > 0) s.drain()
    expect(state.mounted.length).toBe(landed)
    expect(state.loading).toBe(false)
    await expect(waitForFullContent()).resolves.toBeUndefined()
  })

  it('a second phased mount cancels the first tail', () => {
    const blocks = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 3
    const first = makeHost()
    const s1 = makeScheduler()
    setContentPhased(first.host, docOf(blocks), s1.schedule, s1.schedule)
    const second = makeHost()
    const s2 = makeScheduler()
    setContentPhased(second.host, docOf(blocks), s2.schedule, s2.schedule)
    while (s1.pending() > 0) s1.drain()
    while (s2.pending() > 0) s2.drain()
    // the first document keeps only its phase-1 mount; the second is complete
    expect(first.state.mounted.length).toBe(PHASE1_BLOCKS)
    expect(second.state.mounted.length).toBe(blocks)
    expect(first.state.loading).toBe(false)
    expect(second.state.loading).toBe(false)
  })

  it('a rejected chunk falls back to the one-pass mount and releases the save gate', async () => {
    // enough blocks that the adaptive stream needs several appends before the
    // final-chunk fuse collapses the remainder into one chunk
    const blocks = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 8
    const { host, state } = makeHost()
    const { appendNodes, setContent } = host
    let appends = 0
    host.appendNodes = (nodes) => {
      // the whole tail is one fused chunk now — rejecting it must fall back
      if (++appends === 1) throw new Error('schema refused')
      appendNodes(nodes)
    }
    host.setContent = (d) => {
      setContent(d)
      state.dirty = true
    }
    const s = makeScheduler()
    setContentPhased(host, docOf(blocks), s.schedule, s.schedule)
    state.dirty = false
    while (s.pending() > 0) s.drain()
    expect(state.mounted.length).toBe(blocks)
    expect(state.events.at(-3)).toBe(`set:${blocks}`)
    // the remount is not an edit either
    expect(state.dirty).toBe(false)
    expect(state.loading).toBe(false)
    await expect(waitForFullContent()).resolves.toBeUndefined()
  })

  it('reports a pending tail only while chunks remain', () => {
    const { host } = makeHost()
    const s = makeScheduler()
    setContentPhased(host, docOf(PHASED_MIN_BLOCKS), s.schedule, s.schedule)
    expect(isPhasedContentPending()).toBe(false)
    setContentPhased(host, docOf(PHASED_MIN_BLOCKS + 1), s.schedule, s.schedule)
    expect(isPhasedContentPending()).toBe(true)
    while (s.pending() > 0) s.drain()
    expect(isPhasedContentPending()).toBe(false)
    setContentPhased(host, docOf(PHASED_MIN_BLOCKS + 1), s.schedule, s.schedule)
    cancelPhasedContent()
    expect(isPhasedContentPending()).toBe(false)
  })
})

describe('adaptive chunk sizing (PERF-1639)', () => {
  /** scheduler the test drains manually; records each chunk's size */
  function makeRecordingScheduler() {
    const queue: Array<() => void> = []
    return {
      schedule: (cb: () => void) => queue.push(cb),
      drain: () => queue.shift()?.(),
      pending: () => queue.length,
    }
  }

  it('streams the tail in a few large chunks: halves the remainder, then fuses it', () => {
    // 20,000 remaining: 8192 (capped half) -> 5904 (half) -> 5904 (fused remainder)
    const blocks = PHASE1_BLOCKS + 20000
    const { host, state } = makeHost()
    const s = makeRecordingScheduler()
    const sizes: number[] = []
    const baseAppend = host.appendNodes
    host.appendNodes = (nodes) => {
      sizes.push(nodes.length)
      baseAppend(nodes)
    }
    setContentPhased(host, docOf(blocks), s.schedule, s.schedule)
    while (s.pending() > 0) s.drain()
    expect(state.mounted.length).toBe(blocks)
    // every block landed exactly once, in order
    expect(state.mounted.map((n) => n.attrs?.docxIndex)).toEqual(
      Array.from({ length: blocks }, (_, i) => i),
    )
    // first tail chunk stays small (paint escape hatch), then the halve/fuse policy
    expect(sizes).toEqual([PHASE_CHUNK_BLOCKS, PHASE_CHUNK_MAX_BLOCKS, 11680])
  })

  it('keeps a minimum chunk size for small tails and stays within a bounded drain count', () => {
    const blocks = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 8
    const { host, state } = makeHost()
    const s = makeRecordingScheduler()
    setContentPhased(host, docOf(blocks), s.schedule, s.schedule)
    let drains = 0
    while (s.pending() > 0 && drains < 1000) {
      s.drain()
      drains++
    }
    expect(s.pending()).toBe(0)
    expect(state.mounted.length).toBe(blocks)
    // small tail: a single fused chunk
    expect(drains).toBeLessThanOrEqual(2)
  })
})

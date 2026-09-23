import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { MapCache } from '../src/renderer/document/map-cache'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'

const p = (n: number) => `<p class="row">paragraph ${n}</p>`
const doc = (ps: string[]) => `<html><body>${ps.join('')}</body></html>`

describe('MapCache keeps rebuilds off the hot path', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  function makeCache(delay = 300) {
    const onRebuild = vi.fn()
    const build = vi.fn((text: string, version: number, prev: ParseMap | null) =>
      buildParseMap(text, version, prev),
    )
    const cache = new MapCache<ParseMap>(build, onRebuild, delay)
    return { cache, build, onRebuild }
  }

  it('builds synchronously only on the first miss', () => {
    const { cache, build } = makeCache()
    const r1 = cache.get(doc([p(1)]), 0)
    expect(r1.stale).toBe(false)
    expect(build).toHaveBeenCalledTimes(1)
    // unchanged text/version -> cache hit, no extra builds, no pending work
    const r2 = cache.get(doc([p(1)]), 0)
    expect(r2.stale).toBe(false)
    expect(r2.map).toBe(r1.map)
    expect(cache.stats()).toEqual({ builds: 1, staleServes: 0, pending: false })
  })

  it('serves the previous map stale and debounces one rebuild per storm', () => {
    const { cache, build, onRebuild } = makeCache()
    cache.get(doc([p(1)]), 0)
    // typing storm: 10 edits, each bumps version and changes text
    for (let i = 2; i <= 11; i++) {
      const stale = cache.get(doc([p(1), p(i)]), i - 1)
      expect(stale.stale).toBe(true)
      expect(cache.stats().builds).toBe(1) // nothing rebuilt on the hot path
    }
    expect(cache.stats().staleServes).toBe(10)
    expect(cache.stats().pending).toBe(true)
    vi.advanceTimersByTime(300)
    expect(build).toHaveBeenCalledTimes(2) // exactly one rebuild for the whole storm
    expect(onRebuild).toHaveBeenCalledTimes(1)
    const fresh = cache.get(doc([p(1), p(11)]), 10)
    expect(fresh.stale).toBe(false)
    expect(fresh.map.version).toBe(10)
    expect(cache.get(doc([p(1), p(11)]), 10).map).toBe(fresh.map)
  })

  it('reschedules when edits continue past the delay', () => {
    vi.useFakeTimers()
    const { cache, build, onRebuild } = makeCache()
    cache.get(doc([p(1)]), 0)
    vi.advanceTimersByTime(200)
    cache.get(doc([p(1), p(2)]), 1) // re-scheduled before the first fire
    vi.advanceTimersByTime(200)
    expect(build).toHaveBeenCalledTimes(1) // still nothing fired
    cache.get(doc([p(1), p(2), p(3)]), 2)
    vi.advanceTimersByTime(300)
    expect(build).toHaveBeenCalledTimes(2)
    expect(onRebuild).toHaveBeenCalledTimes(1)
  })

  it('now() builds synchronously and cancels the pending rebuild', () => {
    const { cache, build, onRebuild } = makeCache()
    cache.get(doc([p(1)]), 0)
    cache.get(doc([p(1), p(2)]), 1) // stale serve, rebuild pending
    const map = cache.now(doc([p(1), p(2), p(3)]), 2)
    expect(cache.stats().pending).toBe(false)
    expect(cache.stats().builds).toBe(2)
    expect(map.version).toBe(2)
    vi.advanceTimersByTime(1000)
    expect(build).toHaveBeenCalledTimes(2) // no second build from the cancelled timer
    expect(onRebuild).not.toHaveBeenCalled() // now() never fires the rebuild hook
    const again = cache.now(doc([p(1), p(2), p(3)]), 2)
    expect(again).toBe(map)
  })

  it('stays correct after an edit series: fresh map equals a from-scratch build', () => {
    const { cache } = makeCache()
    const series = [
      doc([p(1), p(2), p(3)]),
      doc([p(1), p(2), p(3), p(4)]),
      doc([p(1), '<h2>t</h2>', p(3), p(4)]),
      doc([p(1), '<h2>t</h2>', p(4)]),
    ]
    let text = series[0]
    let version = 0
    const first = cache.now(text, version)
    let prev = first
    for (const next of series) {
      version++
      text = next
      const stale = cache.get(text, version)
      void stale
      const fresh = cache.now(text, version)
      const oracle = buildParseMap(text, version, prev)
      expect(fresh.elements.map((e) => [e.sid, e.range])).toEqual(
        oracle.elements.map((e) => [e.sid, e.range]),
      )
      prev = fresh
    }
  })

  it('documents over the auto-rebuild limit serve stale without ever scheduling', () => {
    const onRebuild = vi.fn()
    const build = vi.fn((text: string, version: number, prev: ParseMap | null) =>
      buildParseMap(text, version, prev),
    )
    // tiny limit stands in for the giant-document policy
    const cache = new MapCache<ParseMap>(build, onRebuild, 300, 64)
    cache.get(doc([p(1)]), 0)
    const stale = cache.get(doc([p(1), p(2), p(3), p(4), p(5)]), 1)
    expect(stale.stale).toBe(true)
    expect(cache.stats()).toEqual({ builds: 1, staleServes: 1, pending: false })
    vi.advanceTimersByTime(5000)
    expect(build).toHaveBeenCalledTimes(1) // no auto rebuild for the giant document
    expect(onRebuild).not.toHaveBeenCalled()
    // on-demand readers still get freshness
    const fresh = cache.now(doc([p(1), p(2), p(3), p(4), p(5)]), 1)
    expect(fresh.version).toBe(1)
    expect(cache.stats().builds).toBe(2)
  })

  it('dispose() drops the pending rebuild without building', () => {
    const { cache, build, onRebuild } = makeCache()
    cache.get(doc([p(1)]), 0)
    cache.get(doc([p(1), p(2)]), 1)
    cache.dispose()
    vi.advanceTimersByTime(5000)
    expect(build).toHaveBeenCalledTimes(1)
    expect(onRebuild).not.toHaveBeenCalled()
    expect(cache.stats().pending).toBe(false)
  })
})

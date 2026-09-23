import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { MapCache } from '../src/renderer/document/map-cache'
import {
  buildParseMap,
  elementCovering,
  sourceText,
  type ParseMap,
} from '../src/renderer/document/parse-map'

/**
 * Regression for the PERF-1648 fix round: commitText bumps the buffer version,
 * so any reader that resolves a sid right after a commit must see the fresh map.
 * The best-effort get() serve still holds the previous offsets — elementCovering
 * against it picks the wrong element and the selection/crumb resets (the e2e
 * html-tab failure where the crumb came back as `section.hero` instead of
 * `p.note` after a toolbar Move).
 */
const here = dirname(fileURLToPath(import.meta.url))
const appSource = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')

/** the exact document the html-tab inspector e2e uses */
const noteLine = '  <p class="note">Second line.</p>\n'
const leadLine = '  <p class="lead">First line.</p>\n'
const source = `<!doctype html>\n<html>\n<body>\n<section class="hero">\n  <h1 id="title">Hello</h1>\n${leadLine}${noteLine}</section>\n</body>\n</html>\n`
/** after Move up: the note line replaced the lead line's slot, shifting every later offset */
const after = source.replace(leadLine + noteLine, noteLine + leadLine)

/** source span of App.tsx between two anchors, so the tripwire fails loudly when refactors move code */
function sliceBetween(haystack: string, startMarker: string, endMarker: string): string {
  const start = haystack.indexOf(startMarker)
  expect(start, `App.tsx no longer declares "${startMarker.trim()}"`).toBeGreaterThanOrEqual(0)
  const end = haystack.indexOf(endMarker, start)
  expect(end, `App.tsx anchor lost: "${endMarker.trim()}"`).toBeGreaterThan(start)
  return haystack.slice(start, end)
}

describe('selection resolves against the fresh map after commitText', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  function makeCache() {
    const onRebuild = vi.fn()
    const cache = new MapCache<ParseMap>((t, v, prev) => buildParseMap(t, v, prev), onRebuild, 300)
    return { cache, onRebuild }
  }

  it('after a commit the reselect target and the caret crumb come from the fresh map', () => {
    const { cache } = makeCache()
    cache.get(source, 0) // initial build, version 0

    // applyOps compiled the move against the pre-commit text and commitText bumped
    // the version; the widest applied range is the inserted note line, trimmed to the element
    const tag = '<p class="note">Second line.</p>'
    const trimmed: [number, number] = [after.indexOf(tag), after.indexOf(tag) + tag.length]

    // what the old code saw: the stale serve holds the pre-move offsets
    const staleServe = cache.get(after, 1)
    expect(staleServe.stale).toBe(true)
    expect(staleServe.map.version).toBe(0)
    // at the new slot the stale map still finds the old structure: the covering
    // element is the structural ancestor, so the reselect reset the selection
    expect(elementCovering(staleServe.map, trimmed[0], trimmed[1])?.tag).toBe('section')

    // what the fixed code sees: the synchronous rebuild matches the committed buffer
    const fresh = cache.now(after, 1)
    expect(fresh.version).toBe(1)
    const reselected = elementCovering(fresh, trimmed[0], trimmed[1])
    expect(reselected?.tag).toBe('p')
    expect(reselected && sourceText(after, reselected)).toBe('Second line.')

    // the crumb follows the same rule: the caret sits inside the moved note, and only
    // the fresh map names it there (the stale map may even serve the same sid for a
    // different element — sid matching by path aliases the sibling that took the slot)
    const caret = after.indexOf('Second line.') + 3
    const freshCrumb = elementCovering(fresh, caret, caret + 1)
    expect(freshCrumb?.tag).toBe('p')
    expect(freshCrumb && sourceText(after, freshCrumb)).toBe('Second line.')
    const staleCrumb = elementCovering(staleServe.map, caret, caret + 1)
    expect(staleCrumb && sourceText(source, staleCrumb)).not.toBe('Second line.')
  })

  it('runManual reselect and onCursor never resolve from the stale best-effort map', () => {
    const runManual = sliceBetween(
      appSource,
      'const runManual = useCallback(',
      'const selectSidRef = useRef<',
    )
    expect(runManual).toContain('const map = getMapNow()')
    expect(runManual).not.toMatch(/getMap\(/)
    expect(runManual).toMatch(/\[applyOps, getMapNow, flushPending\]/)

    const onCursor = sliceBetween(
      appSource,
      'const onCursor = useCallback(',
      'const selectedEntry =',
    )
    // a stale serve must not move the selection onto a wrong element either:
    // onCursor skips resolution until the map has caught up
    expect(onCursor).toContain('getMapState()')
    expect(onCursor).toMatch(/if \(stale\) return/)
    expect(onCursor).not.toMatch(/getMap\(/)
    expect(onCursor).toMatch(/\[getMapState, selectSid\]/)
  })
})

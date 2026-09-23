import { describe, expect, it } from 'vitest'
import { buildParseMap, type ElementEntry, type ParseMap } from '../src/renderer/document/parse-map'

/** the pre-optimization sid matching: linear scan of the previous map in document order */
function buildWithLinearMatch(text: string, version: number, previous: ParseMap | null): ParseMap {
  const used = new Set<number>()
  const linearMatch = (entry: Omit<ElementEntry, 'sid'>): number | null => {
    if (!previous) return null
    let best: ElementEntry | null = null
    let bestDist = Infinity
    for (const old of previous.elements) {
      if (used.has(old.sid) || old.tag !== entry.tag || old.path !== entry.path) continue
      const dist = Math.abs(old.startTag[0] - entry.startTag[0])
      if (dist < bestDist) {
        best = old
        bestDist = dist
      }
    }
    if (!best) return null
    used.add(best.sid)
    return best.sid
  }
  // buildParseMap without previous, then reassign sids linearly in the same order
  const fresh = buildParseMap(text, version, null)
  let next = previous ? Math.max(0, ...previous.elements.map((e) => e.sid)) + 1 : 1
  const byStart = new Map<number, number>()
  const elements: ElementEntry[] = []
  for (const e of fresh.elements) {
    const sid = linearMatch(e) ?? next++
    byStart.set(e.startTag[0], sid)
  }
  // second pass: rebuild entries with the linear sids (start offsets are unique per element)
  const rebased = buildParseMap(text, version, null)
  for (const e of rebased.elements) {
    const sid = byStart.get(e.startTag[0])!
    elements.push({ ...e, sid })
  }
  return {
    version,
    elements,
    bySid: new Map(elements.map((e) => [e.sid, e])),
    errorCount: rebased.errorCount,
  }
}

const doc = (ps: string[]) => `<html><body>${ps.join('')}</body></html>`
const p = (n: number) => `<p class="row">paragraph ${n}</p>`

describe('sid matching across rebuilds', () => {
  it('assigns the same sids as the linear-scan reference over an edit series', () => {
    let text = doc([p(1), p(2), p(3), p(4)])
    let ref = buildWithLinearMatch(text, 1, null)
    let actual = buildParseMap(text, 1, null)
    expect(actual.elements.map((e) => e.sid)).toEqual(ref.elements.map((e) => e.sid))

    const series = [
      // append at the end
      text + '',
      doc([p(1), p(2), p(3), p(4), p(5)]),
      // insert in the middle (paths of later siblings shift by nth-of-type)
      doc([p(1), p(9), p(2), p(3), p(4), p(5)]),
      // delete from the middle
      doc([p(1), p(9), p(3), p(4), p(5)]),
      // replace a tag
      doc([p(1), '<h2>head</h2>', p(3), p(4), p(5)]),
      // duplicate an existing element
      doc([p(1), p(1), '<h2>head</h2>', p(3), p(4), p(5)]),
    ]
    let version = 1
    for (const next of series) {
      version++
      const prevRef = ref
      const prevActual = actual
      text = next
      ref = buildWithLinearMatch(text, version, prevRef)
      actual = buildParseMap(text, version, prevActual)
      expect(actual.elements.map((e) => e.sid)).toEqual(ref.elements.map((e) => e.sid))
      // every sid is unique
      expect(new Set(actual.elements.map((e) => e.sid)).size).toBe(actual.elements.length)
    }
  })

  it('keeps sids stable for untouched elements across a small tail edit', () => {
    const before = doc([p(1), p(2), p(3)])
    const first = buildParseMap(before, 1, null)
    const after = doc([p(1), p(2), p(3), '<p>tail</p>'])
    const second = buildParseMap(after, 2, first)
    const byTag = new Map(second.elements.map((e) => [e.tag === 'p' ? e.inner[0] : -1, e.sid]))
    for (const old of first.elements) {
      if (old.tag !== 'p') continue
      // same source offset -> same sid (only a tail append happened)
      const fresh = byTag.get(old.inner[0])
      expect(fresh).toBe(old.sid)
    }
  })

  it('never reuses one sid twice within a rebuild', () => {
    const text = doc([p(1), p(1), p(1)])
    const first = buildParseMap(text, 1, null)
    const doubled = doc([p(1), p(1), p(1), p(1)])
    const second = buildParseMap(doubled, 2, first)
    const sids = second.elements.map((e) => e.sid)
    expect(new Set(sids).size).toBe(sids.length)
  })
})

import { describe, expect, it } from 'vitest'
import { parse } from 'parse5'
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

describe('nth-of-type paths', () => {
  // the pre-optimization semantics: for every element, walk its parent's
  // childNodes and count same-tag elements up to the node itself
  function referencePaths(text: string): string[] {
    interface El {
      tagName: string
      childNodes?: unknown[]
      sourceCodeLocation?: { startTag?: { startOffset: number } }
    }
    const isEl = (n: unknown): n is El => typeof (n as { tagName?: unknown }).tagName === 'string'
    const tree = parse(text, { sourceCodeLocationInfo: true })
    const out: string[] = []
    const segmentOf = (node: El, siblings: unknown[]): string => {
      if (node.tagName === 'html' || node.tagName === 'head' || node.tagName === 'body')
        return node.tagName
      let n = 0
      for (const sibling of siblings) {
        if (isEl(sibling) && sibling.tagName === node.tagName) {
          n++
          if (sibling === node) return `${node.tagName}:nth-of-type(${n})`
        }
      }
      return `${node.tagName}:nth-of-type(1)`
    }
    const walk = (node: unknown, siblings: unknown[], path: string): void => {
      let own = path
      if (isEl(node)) {
        const loc = node.sourceCodeLocation
        if (loc?.startTag) {
          own = path ? `${path} > ${segmentOf(node, siblings)}` : segmentOf(node, siblings)
          out.push(own)
        }
      }
      const kids = (node as { childNodes?: unknown[] }).childNodes ?? []
      for (const child of kids) walk(child, kids, own)
      // template content lives outside childNodes
      const content = (node as { content?: { childNodes?: unknown[] } }).content
      if (content)
        for (const child of content.childNodes ?? []) walk(child, content.childNodes, own)
    }
    walk(tree, [], '')
    return out
  }

  it('matches the sibling-counting reference on nested and mixed markup', () => {
    const text = `
      <html><body>
        <section><p>one</p><p>two</p><span>x</span><p>three</p></section>
        <section><p>four</p></section>
        <table><tr><td>a</td><td>b</td></tr><tr><td>c</td></tr></table>
        <ul><li>1</li><li>2</li></ul>
        <template><div>t</div><div>t2</div></template>
      </body></html>`
    const first = buildParseMap(text, 1, null)
    expect(first.elements.map((e) => e.path)).toEqual(referencePaths(text))
    expect(first.elements.map((e) => e.path)).toEqual([
      'html',
      'html > body',
      'html > body > section:nth-of-type(1)',
      'html > body > section:nth-of-type(1) > p:nth-of-type(1)',
      'html > body > section:nth-of-type(1) > p:nth-of-type(2)',
      'html > body > section:nth-of-type(1) > span:nth-of-type(1)',
      'html > body > section:nth-of-type(1) > p:nth-of-type(3)',
      'html > body > section:nth-of-type(2)',
      'html > body > section:nth-of-type(2) > p:nth-of-type(1)',
      'html > body > table:nth-of-type(1)',
      'html > body > table:nth-of-type(1) > tr:nth-of-type(1)',
      'html > body > table:nth-of-type(1) > tr:nth-of-type(1) > td:nth-of-type(1)',
      'html > body > table:nth-of-type(1) > tr:nth-of-type(1) > td:nth-of-type(2)',
      'html > body > table:nth-of-type(1) > tr:nth-of-type(2)',
      'html > body > table:nth-of-type(1) > tr:nth-of-type(2) > td:nth-of-type(1)',
      'html > body > ul:nth-of-type(1)',
      'html > body > ul:nth-of-type(1) > li:nth-of-type(1)',
      'html > body > ul:nth-of-type(1) > li:nth-of-type(2)',
      'html > body > template:nth-of-type(1)',
      'html > body > template:nth-of-type(1) > div:nth-of-type(1)',
      'html > body > template:nth-of-type(1) > div:nth-of-type(2)',
    ])
  })
})

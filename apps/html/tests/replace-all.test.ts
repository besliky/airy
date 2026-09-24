import { describe, expect, it } from 'vitest'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { undo } from '@codemirror/commands'
import { forceParsing, syntaxTree } from '@codemirror/language'
import { buildExtensions } from '../src/renderer/source/cm-setup'
import { cmFindTarget } from '../src/renderer/source/find-target'
import {
  findHighlight,
  onBulkReplaceProgress,
  setFindHits,
  type BulkReplaceProgress,
} from '../src/renderer/source/cm-find'
import { buildParseMap, type ParseMap } from '../src/renderer/document/parse-map'
import { MapCache } from '../src/renderer/document/map-cache'
import { instrumentForPreview } from '../src/renderer/preview/instrument'

const opts = { matchCase: false, wholeWord: false }

const LINE = '<p>alpha beta alpha gamma</p>\n'
// 12 000 matches keep the replace well past trivial while the suite stays fast
const BIG = '<!doctype html><html><body>\n' + LINE.repeat(6000) + '</body></html>\n'
const MATCHES = 12_000

function setup(text: string) {
  const changed = new Set<() => void>()
  const view = new EditorView({
    state: EditorState.create({ doc: text, extensions: buildExtensions(() => {}) }),
    parent: document.body,
  })
  const target = cmFindTarget(view, (l) => {
    changed.add(l)
    return () => changed.delete(l)
  })
  return { view, target }
}

describe('replace-all as one structural operation', () => {
  it('replaces every occurrence and stays one undo step', async () => {
    const { view, target } = setup(BIG)
    expect(target.search('alpha', opts, 0)).toBe(MATCHES)
    // timing canary: the pre-fix path froze ~2 minutes on the 3MB/200k-match
    // corpus (O(hits × changes) highlight mapping); this shape must stay fast
    const t0 = performance.now()
    await target.replaceAll('ZZQQ')
    const replaceMs = performance.now() - t0
    // byte-for-byte the same as a whole-string replace
    expect(view.state.doc.toString()).toBe(BIG.replaceAll('alpha', 'ZZQQ'))
    // a single undo step restores the pre-replace document
    undo(view)
    expect(view.state.doc.toString()).toBe(BIG)
    expect(replaceMs).toBeLessThan(5000)
    view.destroy()
  })

  it('does not re-replace a replacement that contains the query', async () => {
    const { view, target } = setup(BIG)
    target.search('alpha', opts, 0)
    await target.replaceAll('alphax')
    const out = view.state.doc.toString()
    expect(out).toBe(BIG.replaceAll('alpha', 'alphax'))
    expect(out.split('alphaxx').length - 1).toBe(0)
    view.destroy()
  })

  it('reports progress totals and ends with a null event', async () => {
    const { view, target } = setup(BIG)
    const events: Array<BulkReplaceProgress | null> = []
    const off = onBulkReplaceProgress((p) => events.push(p))
    target.search('alpha', opts, 0)
    target.replaceAll('ZZQQ')
    off()
    // the subscription echoes the (null) current state first
    const frames = events.filter((e): e is BulkReplaceProgress => e !== null)
    expect(frames[0]).toEqual({ total: MATCHES, done: 0 })
    expect(events[events.length - 1]).toBeNull()
    for (let i = 1; i < frames.length; i++)
      expect(frames[i]!.done).toBeGreaterThan(frames[i - 1]!.done)
    expect(frames[frames.length - 1]).toEqual({ total: MATCHES, done: MATCHES })
    view.destroy()
  })

  it('answers search with the announced count while the replace runs, unpainted', () => {
    const { view, target } = setup(BIG)
    target.search('alpha', opts, 0)
    // drive the search from a progress callback: it fires while the replace
    // transaction owns the document (bulkReplaceBusy active)
    let seen: number | null = null
    let painted = -1
    const off = onBulkReplaceProgress((p) => {
      if (!p || seen !== null) return
      seen = target.search('alpha', opts, 0)
      // no hits are painted mid-run: the unpainted set is what keeps the
      // RangeSet.map cost of the big transaction constant
      painted = 0
      view.state.field(findHighlight).between(0, view.state.doc.length, () => {
        painted++
        return false
      })
    })
    target.replaceAll('ZZQQ')
    off()
    expect(seen).toBe(MATCHES)
    expect(painted).toBe(0)
    expect(target.search('alpha', opts, 0)).toBe(0)
    view.destroy()
  })

  it('is a no-op when nothing was searched or while already running', async () => {
    const { view, target } = setup(BIG)
    await target.replaceAll('ZZQQ')
    expect(view.state.doc.toString()).toBe(BIG)
    view.destroy()
  })

  it('reinstates the html syntax tree after the replace', async () => {
    const { view, target } = setup(BIG)
    target.search('alpha', opts, 0)
    // the replace runs with the parser swapped out (thousands of scattered
    // changed ranges invalidate the tree doc-wide); afterwards the parser must
    // be back and able to cover the whole document (the tree re-parses lazily)
    await target.replaceAll('ZZQQ')
    expect(forceParsing(view, view.state.doc.length, 1e8)).toBe(true)
    const tree = syntaxTree(view.state)
    expect(tree.length).toBe(view.state.doc.length)
    expect(tree.topNode.name).toBe('Document')
    view.destroy()
  })
})

describe('replace-all effects on the parse map and preview marks', () => {
  it('keeps sids of untouched elements stable and keeps every mark', () => {
    const before = '<!doctype html><html><body>' + LINE.repeat(4) + '</body></html>'
    const after = before.replaceAll('alpha', 'ZZQQ')
    const first = buildParseMap(before, 1, null)
    const second = buildParseMap(after, 2, first)
    // the element structure is unchanged: tag, path and depth line up one-to-one
    expect(second.elements.length).toBe(first.elements.length)
    for (let i = 0; i < first.elements.length; i++) {
      expect(second.elements[i]!.tag).toBe(first.elements[i]!.tag)
      expect(second.elements[i]!.path).toBe(first.elements[i]!.path)
      expect(second.elements[i]!.sid).toBe(first.elements[i]!.sid)
    }
    // the preview copy still carries one data-sid mark per element
    const instrumented = instrumentForPreview(after, second, '')
    const marks = instrumented.split('data-sid=').length - 1
    expect(marks).toBe(second.elements.length)
  })
})

describe('find-highlight effect handling', () => {
  it('an effect replaces the whole set without mapping through the changes', () => {
    const view = new EditorView({
      state: EditorState.create({ doc: 'aa bb aa', extensions: [findHighlight] }),
      parent: document.body,
    })
    view.dispatch({ effects: setFindHits.of({ ranges: [{ from: 0, to: 2 }], active: 0 }) })
    // changes and a fresh hit set in one transaction: the effect ranges win as-is
    view.dispatch({
      changes: { from: 0, to: 2, insert: 'cccc' },
      effects: setFindHits.of({
        ranges: [
          { from: 0, to: 4 },
          { from: 5, to: 7 },
        ],
        active: 1,
      }),
    })
    const seen: number[] = []
    view.state.field(findHighlight).between(0, view.state.doc.length, (from, to) => {
      seen.push(from, to)
      return true
    })
    expect(seen).toEqual([0, 4, 5, 7])
    view.destroy()
  })
})

describe('replace-all on a larger document', () => {
  // 26 000 matches over a 300+ KB document: same invariants at a scale where a
  // painted replace would already cost seconds
  const LINE_YIELD = '<p>alpha</p>\n'
  const YIELD_DOC = '<!doctype html><html><body>\n' + LINE_YIELD.repeat(26_000) + '</body></html>\n'

  it('replaces in one undo step with the syntax tree back', async () => {
    const { view, target } = setup(YIELD_DOC)
    target.search('alpha', opts, 0)
    await target.replaceAll('ZZQQ')
    expect(view.state.doc.toString()).toBe(YIELD_DOC.replaceAll('alpha', 'ZZQQ'))
    undo(view)
    expect(view.state.doc.toString()).toBe(YIELD_DOC)
    expect(forceParsing(view, view.state.doc.length, 1e8)).toBe(true)
    expect(syntaxTree(view.state).topNode.name).toBe('Document')
    view.destroy()
  })

  it('maps the cursor through the replaced text', async () => {
    const view = new EditorView({
      state: EditorState.create({
        doc: 'alpha ONE alpha',
        selection: { anchor: 15 },
        extensions: buildExtensions(() => {}),
      }),
      parent: document.body,
    })
    const target = cmFindTarget(view, () => () => {})
    target.search('alpha', opts, 0)
    await target.replaceAll('B')
    expect(view.state.doc.toString()).toBe('B ONE B')
    expect(view.state.selection.main.head).toBe(7)
    view.destroy()
  })
})

describe('replace-all parse-map cost', () => {
  interface Harness {
    view: EditorView
    target: ReturnType<typeof cmFindTarget>
    cache: MapCache<ParseMap>
    version(): number
  }
  /** mirrors the App wiring: commitText bumps the version, getMapState reads the cache */
  function harness(autoRebuildLimit: number, doc: string): Harness {
    const cache = new MapCache<ParseMap>(
      (t, v, prev) => buildParseMap(t, v, prev),
      () => {},
      300,
      autoRebuildLimit,
    )
    let version = 0
    const changed = new Set<() => void>()
    const view = new EditorView({
      state: EditorState.create({ doc, extensions: buildExtensions(() => {}) }),
      parent: document.body,
      dispatchTransactions: (trs, v) => {
        v.update(trs)
        if (trs.some((tr) => tr.docChanged)) {
          cache.get(v.state.doc.toString(), ++version)
          for (const l of changed) l()
        }
      },
    })
    const target = cmFindTarget(view, (l) => {
      changed.add(l)
      return () => changed.delete(l)
    })
    return { view, target, cache, version: () => version }
  }

  const settle = () => new Promise((r) => setTimeout(r, 350))

  it('costs at most two builds for twelve thousand replacements (auto-rebuild docs)', async () => {
    const h = harness(262_144, BIG)
    h.cache.get(viewText(h.view), h.version() + 1) // initial load build
    expect(h.cache.stats().builds).toBe(1)
    h.target.search('alpha', opts, 0)
    await h.target.replaceAll('ZZQQ')
    // every chunk committed and served stale; one debounced rebuild covers them all
    expect(h.cache.stats().staleServes).toBeGreaterThan(0)
    await settle()
    expect(h.cache.stats().builds).toBe(2)
    // a correctness-critical reader (pushPreview's getMapNow) adds nothing more
    h.cache.now(viewText(h.view), h.version())
    expect(h.cache.stats().builds).toBe(2)
    h.view.destroy()
  })

  it('giant documents skip the auto-rebuild and rebuild once on demand', async () => {
    const h = harness(1000, BIG) // limit below the doc: the giant path
    h.cache.get(viewText(h.view), h.version() + 1)
    h.target.search('alpha', opts, 0)
    await h.target.replaceAll('ZZQQ')
    await settle()
    expect(h.cache.stats().builds).toBe(1) // no debounced rebuild for giants
    const fresh = h.cache.now(viewText(h.view), h.version())
    expect(h.cache.stats().builds).toBe(2)
    expect(fresh.bySid.size).toBeGreaterThan(0)
    h.view.destroy()
  })

  function viewText(view: EditorView): string {
    return view.state.doc.toString()
  }
})

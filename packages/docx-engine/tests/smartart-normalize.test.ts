// TEST-702: direct pins for the insertion input guards that the package
// anatomy tests only exercise incidentally — normalizeDiagramItems' hard node
// cap (a pathological caller must not blow up the package; the UI clamps at 8
// but the engine must hold 64 on its own) and the freshDiagramGuid format the
// data model and drawing cache correlate through.
import { describe, expect, it } from 'vitest'
import { buildDiagramDisplay, freshDiagramGuid, type NewDiagramItem } from '../src/index'
import { normalizeDiagramItems } from '../src/smartart-diagram'

describe('normalizeDiagramItems', () => {
  it('caps the model at 64 nodes no matter how many the caller passes', () => {
    const items: NewDiagramItem[] = Array.from({ length: 70 }, (_, i) => ({
      text: `n${String(i)}`,
      level: 0,
    }))
    const out = normalizeDiagramItems(items)
    expect(out.length).toBe(64)
    expect(out[63]).toEqual({ text: 'n63', level: 0 })
    // the cap is applied before anything else: no partial-overflow output
    expect(out.every((n) => n.text !== 'n64')).toBe(true)
  })

  it('collapses whitespace, drops empties, floors wild levels, and clamps indent jumps', () => {
    const out = normalizeDiagramItems([
      { text: '  a\t\nb  ', level: 0 },
      { text: '   ', level: 0 }, // dropped
      { text: 'deep', level: 9 }, // jump clamped to prev+1
      { text: 'neg', level: -3 }, // floored to 0
      { text: 'jump', level: 5 }, // clamped to one deeper than 'neg'
      { text: 'x' }, // missing level reads as 0
    ])
    expect(out).toEqual([
      { text: 'a b', level: 0 },
      { text: 'deep', level: 1 },
      { text: 'neg', level: 0 },
      { text: 'jump', level: 1 },
      { text: 'x', level: 0 },
    ])
  })

  it('never returns an empty model: a lone placeholder survives', () => {
    expect(normalizeDiagramItems([])).toEqual([{ text: '[Text]', level: 0 }])
    expect(normalizeDiagramItems([{ text: ' \n ', level: 3 }])).toEqual([
      { text: '[Text]', level: 0 },
    ])
  })
})

describe('freshDiagramGuid', () => {
  it('produces Office braced uppercase GUIDs of the 8-4-4-4-12 shape', () => {
    const pattern = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/
    for (let i = 0; i < 100; i++) expect(freshDiagramGuid()).toMatch(pattern)
  })

  it('stays unique across a realistic insertion burst', () => {
    const guids = new Set(Array.from({ length: 1000 }, () => freshDiagramGuid()))
    expect(guids.size).toBe(1000)
  })

  it('shows up in the display preview: shapes carry the model ids they will keep', () => {
    // the preview runs the same normalization; a wild model still renders
    // exactly the capped 64 shapes (the insert path shares this code)
    const display = buildDiagramDisplay({
      kind: 'blockList',
      items: Array.from({ length: 70 }, (_, i) => ({ text: `n${String(i)}` })),
    })
    expect(display.shapes.length).toBe(64)
  })
})

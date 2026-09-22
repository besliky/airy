/**
 * Groups containing tables (PAR-316): a p:graphicFrame table groups together
 * with shapes like any editable element — the table rides along as a group
 * child (parse and render layers already handled tables inside groups) and
 * comes back intact on ungroup, with its p:xfrm mapped back to slide
 * coordinates. Regression gate for the GROUPABLE widening.
 */
import { describe, it, expect } from 'vitest'
import {
  addElement,
  addTable,
  createBlankPptx,
  groupElements,
  openPptx,
  savePptx,
  ungroupElement,
} from '../src/index'
import type { GroupElement, TableElement } from '../src/types'

async function setup() {
  const opened = await openPptx(await createBlankPptx())
  addElement(opened.deck.slides[0]!, {
    kind: 'rect',
    offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
    fillColor: '#4472C4',
  })
  const tbl = addTable(opened, 0, {
    rows: 2,
    cols: 2,
    offset: { x: 914400, y: 2743200, cx: 3657600, cy: 1828800 },
  })!
  // addTable materializes the slide: element ids re-resolve from the fresh model
  const elements = opened.deck.slides[0]!.elements
  const spId = elements.find((e) => e.type === 'shape')!.id
  return { opened, spId, tblId: tbl.elementId }
}

describe('groupElements with tables', () => {
  it('groups a shape and a table; the group children keep the table', async () => {
    const { opened, spId, tblId } = await setup()
    const before = opened.deck.slides[0]!.elements.find((e) => e.id === tblId) as TableElement
    const result = groupElements(opened, 0, [spId, tblId])
    expect(result).not.toBeNull()
    const grp = result!.slide.elements.find((e) => e.id === result!.groupId) as GroupElement
    expect(grp).toBeDefined()
    expect(grp.type).toBe('group')
    const childTypes = grp.children.map((c) => c.type).sort()
    expect(childTypes).toEqual(['shape', 'table'])
    const tblChild = grp.children.find((c) => c.type === 'table') as TableElement
    // The table keeps its grid through the group roundtrip
    expect(tblChild.colWidths).toEqual(before.colWidths)
    expect(tblChild.rowHeights).toEqual(before.rowHeights)
    // chOff/chExt are 1:1 with the bbox: child coords stay slide coordinates
    expect(grp.childOffset?.x).toBe(grp.transform.offset.x)
    expect(grp.childOffset?.y).toBe(grp.transform.offset.y)
  })

  it('group + save + reopen: the table survives inside the group', async () => {
    const { opened, spId, tblId } = await setup()
    const result = groupElements(opened, 0, [spId, tblId])
    expect(result).not.toBeNull()
    const reopened = await openPptx(await savePptx(opened))
    const grp = reopened.deck.slides[0]!.elements.find((e) => e.type === 'group') as GroupElement
    expect(grp).toBeDefined()
    const tblChild = grp.children.find((c) => c.type === 'table') as TableElement | undefined
    expect(tblChild).toBeDefined()
    expect(tblChild!.rows.length).toBe(2)
    // Raw bytes still carry the graphicFrame table
    expect(grp.anchor.originalXml).toContain('<a:tbl>')
  })

  it('ungroup restores the table at top level with slide coordinates', async () => {
    const { opened, spId, tblId } = await setup()
    const before = opened.deck.slides[0]!.elements.find((e) => e.id === tblId) as TableElement
    const result = groupElements(opened, 0, [spId, tblId])
    expect(result).not.toBeNull()

    const ungrouped = ungroupElement(opened, 0, result!.groupId)
    expect(ungrouped).not.toBeNull()
    const tbl = ungrouped!.elements.find((e) => e.type === 'table') as TableElement | undefined
    expect(tbl).toBeDefined()
    // 1:1 child coordinate system → the table lands back on its original rect
    expect(tbl!.transform.offset).toEqual(before.transform.offset)
    expect(tbl!.colWidths).toEqual(before.colWidths)

    const out = await savePptx(opened)
    const reopened = await openPptx(out)
    const fresh = reopened.deck.slides[0]!.elements
    expect(fresh.some((e) => e.type === 'group')).toBe(false)
    const freshTbl = fresh.find((e) => e.type === 'table') as TableElement | undefined
    expect(freshTbl).toBeDefined()
    expect(freshTbl!.rows.length).toBe(2)
    expect(freshTbl!.transform.offset.x).toBe(before.transform.offset.x)
  })

  it('refuses grouping a table with a chart passthrough (allowlist still enforced)', async () => {
    const opened = await openPptx(await createBlankPptx())
    const tbl = addTable(opened, 0, {
      rows: 1,
      cols: 1,
      offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    })!
    // A passthrough element would be refused outright; synthesize one the way
    // appendRawElements does before materialization (kind 'unknown').
    opened.deck.slides[0]!.elements.push({
      id: 'raw_pt',
      type: 'passthrough',
      kind: 'unknown',
      anchor: { spIndex: -1, originalXml: '<p:sp>unknown</p:sp>', range: [0, 0] },
      transform: { offset: { x: 0, y: 0, cx: 10, cy: 10 }, rot: 0, flipH: false, flipV: false },
    } as never)
    const result = groupElements(opened, 0, [tbl.elementId, 'raw_pt'])
    expect(result).toBeNull()
  })
})

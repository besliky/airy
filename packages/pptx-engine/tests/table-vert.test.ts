/**
 * Table cell text direction (tcPr@vert, BUG-1673): the parser maps the attribute onto
 * the cell text body so vertical header cells render rotated, and save keeps the
 * attribute (pure presentation defect — the data was never lost; these tests pin it).
 */
import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { addTable, createBlankPptx, editTableCellText, openPptx, savePptx } from '../src/index'
import type { TableElement } from '../src/types'

const SLIDE_XML = 'ppt/slides/slide1.xml'

/** Blank deck + a 2x2 table whose every tcPr carries vert="vert270" (injected via zip patch) */
async function deckWithVertCells(): Promise<Buffer> {
  const opened = await openPptx(await createBlankPptx())
  addTable(opened, 0, {
    rows: 2,
    cols: 2,
    offset: { x: 914400, y: 914400, cx: 4572000, cy: 1828800 },
  })
  const zip = await JSZip.loadAsync(await savePptx(opened))
  const xml = await zip.file(SLIDE_XML)!.async('string')
  // addTable writes plain <a:tcPr/> per cell
  zip.file(SLIDE_XML, xml.replaceAll('<a:tcPr/>', '<a:tcPr vert="vert270"/>'))
  return zip.generateAsync({ type: 'nodebuffer' })
}

const tableOf = (deck: Awaited<ReturnType<typeof openPptx>>['deck']) =>
  deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement

describe('table tcPr@vert parse + save round trip', () => {
  it('parses tcPr@vert onto the cell text body (the render-facing model field)', async () => {
    const { deck } = await openPptx(await deckWithVertCells())
    const tbl = tableOf(deck)
    for (const row of tbl.rows) {
      for (const cell of row) {
        expect(cell.merged ?? false).toBe(false)
        expect(cell.text?.vert).toBe('vert270')
      }
    }
  })

  it('save keeps the vert attribute untouched (no edits → original bytes)', async () => {
    const opened = await openPptx(await deckWithVertCells())
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const xml = await zip.file(SLIDE_XML)!.async('string')
    expect(xml).toContain('<a:tcPr vert="vert270"/>')
    // And it still parses back as vertical text
    const reopened = await openPptx(await savePptx(opened))
    expect(tableOf(reopened.deck).rows[0]![0]!.text?.vert).toBe('vert270')
  })

  it('editing a vert cell keeps its tcPr vert attribute (txBody-only byte surgery)', async () => {
    const opened = await openPptx(await deckWithVertCells())
    const tbl = tableOf(opened.deck)
    expect(
      editTableCellText(opened.deck.slides[0]!, tbl.id, 0, 0, [{ runs: [{ text: 'New' }] }]),
    ).toBe(true)
    const saved = await savePptx(opened)
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file(SLIDE_XML)!.async('string')
    expect(xml).toContain('New')
    expect(xml).toContain('<a:tcPr vert="vert270"/>')
    const reopened = await openPptx(saved)
    const cell = tableOf(reopened.deck).rows[0]![0]!
    expect(cell.text?.vert).toBe('vert270')
    expect(cell.text?.paragraphs[0]!.runs[0]!.text).toBe('New')
  })
})

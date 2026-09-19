// Minimal .pptx fixture builder for the slides session tests: the engine's
// own blank deck (16:9, one empty slide) plus two text-bearing elements,
// serialized through the engine so the bytes are a real package the parser
// round-trips.
import { addElement, createBlankPptx, openPptx, savePptx } from '@airy-office/pptx-engine'

/** one-slide 16:9 deck: a title-ish text box and a filled autoshape with text */
export async function buildFixturePptx(): Promise<Uint8Array> {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  addElement(slide, {
    kind: 'textbox',
    offset: { x: 914_400, y: 365_760, cx: 9_144_000, cy: 1_097_280 },
    paragraphs: [{ runs: [{ text: 'Quarterly Review', bold: true }] }],
  })
  addElement(slide, {
    kind: 'roundRect',
    offset: { x: 914_400, y: 1_828_800, cx: 4_572_000, cy: 1_828_800 },
    fillColor: '#C43E1C',
    paragraphs: [{ runs: [{ text: 'Revenue grew by 12 percent' }] }],
  })
  return savePptx(opened)
}

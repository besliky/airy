/** setElementTextBodyProps: bodyPr direction/autofit/insets/wrap/columns byte surgery. */
import { describe, it, expect } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  savePptx,
  setElementTextBodyProps,
} from '../src/index'
import type { TextElement } from '../src/types'

async function textboxSlide() {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const el = addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
    paragraphs: [{ runs: [{ text: 'x' }] }],
  })
  return { slide, el: el as TextElement }
}

describe('setElementTextBodyProps', () => {
  it('vert attribute set and cleared, model kept in sync', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { vert: 'eaVert' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr vert="eaVert"/)
    expect(el.text!.vert).toBe('eaVert')
    expect(setElementTextBodyProps(slide, el.id, { vert: 'horz' })).toBe(true)
    expect(el.anchor.originalXml).not.toMatch(/<a:bodyPr[^>]*\svert="/)
    expect(el.text!.vert).toBeUndefined()
  })

  it('wrap toggles wrap="none" / wrap="square"', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { wrap: false })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr wrap="none"/)
    expect(el.text!.wrap).toBe(false)
    expect(setElementTextBodyProps(slide, el.id, { wrap: true })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr wrap="square"/)
    expect(el.text!.wrap).toBe(true)
  })

  it('insets written per side (EMU), previously written sides kept', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { insets: { l: 180000 } })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/lIns="180000"/)
    expect(el.text!.insets).toMatchObject({ l: 180000, t: 45720 })
    expect(setElementTextBodyProps(slide, el.id, { insets: { t: 0 } })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/tIns="0"/)
    expect(el.anchor.originalXml).toMatch(/lIns="180000"/)
    expect(el.text!.insets).toMatchObject({ l: 180000, t: 0 })
  })

  it('autofit child swapped in place (self-closing bodyPr expands)', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'shrink' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/)
    expect(el.text!.autofit).toBe('shrink')
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'resize' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:spAutoFit\/>/)
    expect(el.anchor.originalXml).not.toMatch(/normAutofit/)
    expect(el.text!.autofit).toBe('resize')
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'none' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:noAutofit\/>/)
    expect(el.text!.autofit).toBe('none')
  })

  it('numCol/spcCol written as attributes and kept in sync with the model', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { numCol: 3, spcCol: 457200 })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr[^>]*numCol="3"/)
    expect(el.anchor.originalXml).toMatch(/spcCol="457200"/)
    expect(el.text!.numCol).toBe(3)
    expect(el.text!.spcCol).toBe(457200)
    // back to one column: both attributes drop (PowerPoint drops the gap too)
    expect(setElementTextBodyProps(slide, el.id, { numCol: 1 })).toBe(true)
    expect(el.anchor.originalXml).not.toMatch(/numCol="/)
    expect(el.anchor.originalXml).not.toMatch(/spcCol="/)
    expect(el.text!.numCol).toBeUndefined()
    expect(el.text!.spcCol).toBeUndefined()
    // a zero gap removes just spcCol
    expect(setElementTextBodyProps(slide, el.id, { numCol: 2, spcCol: 457200 })).toBe(true)
    expect(setElementTextBodyProps(slide, el.id, { spcCol: 0 })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/numCol="2"/)
    expect(el.anchor.originalXml).not.toMatch(/spcCol="/)
    expect(el.text!.numCol).toBe(2)
    expect(el.text!.spcCol).toBeUndefined()
  })

  it('columns survive a save/reopen round-trip', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
      paragraphs: [{ runs: [{ text: 'x' }] }],
    }) as TextElement
    expect(setElementTextBodyProps(slide, el.id, { numCol: 2, spcCol: 457200 })).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const rel = reopened.deck.slides[0]!.elements[0] as TextElement
    expect(rel.text!.numCol).toBe(2)
    expect(rel.text!.spcCol).toBe(457200)
  })

  it('combined patch applies attributes and autofit together', async () => {
    const { slide, el } = await textboxSlide()
    expect(
      setElementTextBodyProps(slide, el.id, {
        vert: 'vert270',
        wrap: false,
        insets: { b: 91440 },
        autofit: 'shrink',
      }),
    ).toBe(true)
    const xml = el.anchor.originalXml
    expect(xml).toMatch(/vert="vert270"/)
    expect(xml).toMatch(/wrap="none"/)
    expect(xml).toMatch(/bIns="91440"/)
    expect(xml).toMatch(/<a:normAutofit\/>/)
  })

  it('warp writes/replaces/removes <a:prstTxWarp> in schema position', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { warp: { prst: 'textArchUp' } })).toBe(true)
    expect(el.anchor.originalXml).toMatch(
      /<a:bodyPr[^>]*><a:prstTxWarp prst="textArchUp"><a:avLst\/><\/a:prstTxWarp>/,
    )
    expect(el.text!.txWarp).toEqual({ prst: 'textArchUp' })
    // replacing keeps one element; adj values land in avLst
    expect(
      setElementTextBodyProps(slide, el.id, {
        warp: { prst: 'textCircle', adj: { adj: 25000 } },
      }),
    ).toBe(true)
    expect(el.anchor.originalXml).toMatch(
      /<a:prstTxWarp prst="textCircle"><a:avLst><a:gd name="adj" fmla="val 25000"\/><\/a:avLst><\/a:prstTxWarp>/,
    )
    expect((el.anchor.originalXml.match(/<a:prstTxWarp/g) ?? []).length).toBe(1)
    expect(el.text!.txWarp).toEqual({ prst: 'textCircle', adj: { adj: 25000 } })
    // autofit lands after the warp (schema order)
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'shrink' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/prstTxWarp><a:normAutofit\/>/)
    expect(setElementTextBodyProps(slide, el.id, { warp: null })).toBe(true)
    expect(el.anchor.originalXml).not.toMatch(/prstTxWarp/)
    expect(el.text!.txWarp).toBeUndefined()
  })

  it('warp survives a save/reopen round-trip', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
      paragraphs: [{ runs: [{ text: 'x' }] }],
    }) as TextElement
    expect(
      setElementTextBodyProps(slide, el.id, {
        warp: { prst: 'textButton', adj: { adj: 12500 } },
      }),
    ).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const rel = reopened.deck.slides[0]!.elements[0] as TextElement
    expect(rel.text!.txWarp).toEqual({ prst: 'textButton', adj: { adj: 12500 } })
  })

  it('rejects elements without a text body', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    expect(setElementTextBodyProps(slide, 'nope', { wrap: false })).toBe(false)
  })
})

/**
 * Notes formatting (PAR-316): setSlideNotes applies a whole-body format
 * (bold/italic/font size) to every run's a:rPr; getSlideNotesFormat reads it
 * back; plain-text content and the default path (no format) stay unchanged.
 */
import { describe, it, expect } from 'vitest'
import {
  createBlankPptx,
  getSlideNotes,
  getSlideNotesFormat,
  notesPathForSlide,
  openPptx,
  savePptx,
  setSlideNotes,
} from '../src/index'

/** The notesSlide part path of slide 0 (created by setSlideNotes above). */
function notesPathOf(opened: Awaited<ReturnType<typeof openPptx>>): string {
  const p = notesPathForSlide(opened.archive, opened.deck.slides[0]!.path)
  if (!p) throw new Error('notesSlide part not created')
  return p
}

describe('notes formatting', () => {
  it('writes sz/b/i run properties when a format is given', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(
      setSlideNotes(opened, 0, 'big line\nsecond', { bold: true, italic: true, fontSizePt: 28 }),
    ).toBe(true)
    const xml = opened.archive.readText(
      // the notes part was just created for slide 0
      notesPathOf(opened),
    )
    expect(xml).toContain('sz="2800"')
    expect(xml).toContain('b="1"')
    expect(xml).toContain('i="1"')
    // plain text roundtrip is untouched
    expect(getSlideNotes(opened.archive, opened.deck.slides[0]!.path)).toBe('big line\nsecond')
  })

  it('getSlideNotesFormat reads the format back; no format → null', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(getSlideNotesFormat(opened.archive, opened.deck.slides[0]!.path)).toBeNull()
    setSlideNotes(opened, 0, 'styled', { fontSizePt: 14 })
    expect(getSlideNotesFormat(opened.archive, opened.deck.slides[0]!.path)).toEqual({
      fontSizePt: 14,
    })
    setSlideNotes(opened, 0, 'bold only', { bold: true })
    expect(getSlideNotesFormat(opened.archive, opened.deck.slides[0]!.path)).toEqual({ bold: true })
    // rewriting without a format resets to the defaults
    setSlideNotes(opened, 0, 'plain again')
    expect(getSlideNotesFormat(opened.archive, opened.deck.slides[0]!.path)).toBeNull()
  })

  it('format survives save → reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'deck notes', { bold: true, fontSizePt: 18 })
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe('deck notes')
    expect(getSlideNotesFormat(reopened.archive, reopened.deck.slides[0]!.path)).toEqual({
      bold: true,
      fontSizePt: 18,
    })
  })

  it('empty body paragraphs carry the format on endParaRPr', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, '\n', { fontSizePt: 20 })
    const xml = opened.archive.readText(notesPathOf(opened))
    expect(xml).toContain('<a:endParaRPr lang="zh-CN" sz="2000"/>')
    expect(getSlideNotesFormat(opened.archive, opened.deck.slides[0]!.path)).toEqual({
      fontSizePt: 20,
    })
  })
})

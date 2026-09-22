import { describe, expect, it } from 'vitest'
import { nextNotesFormat, notesPtOr } from '../src/renderer/notes-format'

/**
 * PAR-316 notes formatting: the toolbar's whole-body format patches merge
 * like PowerPoint's select-all toggles — explicit booleans flip on/off, size
 * steps clamp to 8..96pt and reverting to the default drops the explicit sz.
 */
describe('nextNotesFormat', () => {
  it('starts from no explicit format', () => {
    expect(nextNotesFormat({}, {})).toEqual({})
    expect(notesPtOr({})).toBe(12)
  })

  it('toggles bold and italic on and back off', () => {
    let fmt = nextNotesFormat({}, { bold: true })
    expect(fmt).toEqual({ bold: true })
    fmt = nextNotesFormat(fmt, { italic: true })
    expect(fmt).toEqual({ bold: true, italic: true })
    fmt = nextNotesFormat(fmt, { bold: false })
    expect(fmt).toEqual({ italic: true })
    fmt = nextNotesFormat(fmt, { italic: false })
    expect(fmt).toEqual({})
  })

  it('steps the font size and clamps to the practical range', () => {
    let fmt = nextNotesFormat({}, { sizeDelta: 1 })
    expect(fmt.fontSizePt).toBe(13)
    for (let i = 0; i < 200; i++) fmt = nextNotesFormat(fmt, { sizeDelta: 1 })
    expect(fmt.fontSizePt).toBe(96)
    for (let i = 0; i < 300; i++) fmt = nextNotesFormat(fmt, { sizeDelta: -1 })
    expect(fmt.fontSizePt).toBe(8)
  })

  it('drops the explicit size again when stepped back to the default', () => {
    let fmt = nextNotesFormat({ fontSizePt: 13 }, { sizeDelta: -1 })
    expect(fmt).toEqual({})
    // an untouched explicit size survives other patches
    fmt = nextNotesFormat({ fontSizePt: 28, bold: true }, { italic: true })
    expect(fmt).toEqual({ bold: true, italic: true, fontSizePt: 28 })
  })

  it('keeps an existing size when only toggling style', () => {
    expect(notesPtOr(nextNotesFormat({ fontSizePt: 20 }, { bold: true }))).toBe(20)
  })
})

import type { NoteNumbering } from '@airy-office/docx-engine'

/** Word's default endnote numbering is lowercase roman (footnotes stay arabic) — the visual cue that separates the two note kinds. */
const ROMAN: Array<[number, string]> = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
]

export function toRoman(n: number): string {
  if (!Number.isFinite(n) || n < 1) return String(n)
  let rest = Math.floor(n)
  let out = ''
  for (const [value, glyph] of ROMAN) {
    while (rest >= value) {
      out += glyph
      rest -= value
    }
  }
  return out
}

const LETTERS = 'abcdefghijklmnopqrstuvwxyz'

/** 1→a, 27→aa (Word's letter numbering wraps like this) */
export function toLetter(n: number, upper = false): string {
  if (!Number.isFinite(n) || n < 1) return String(n)
  let rest = Math.floor(n)
  let out = ''
  do {
    rest -= 1
    out = LETTERS[rest % 26]! + out
    rest = Math.floor(rest / 26)
  } while (rest > 0)
  return upper ? out.toUpperCase() : out
}

/** one value in the numbering formats Word's note dialog offers */
export function formatNoteMark(fmt: NoteNumbering['numFmt'], no: number): string {
  switch (fmt) {
    case 'lowerLetter':
      return toLetter(no)
    case 'upperLetter':
      return toLetter(no, true)
    case 'lowerRoman':
      return toRoman(no)
    case 'upperRoman':
      return toRoman(no).toUpperCase()
    default:
      return String(no)
  }
}

/**
 * Marker of note #no under its kind's options: a custom mark replaces every
 * number (Word's w:customMarkFollows), numStart shifts the first value.
 */
export function noteMarkerOf(opts: NoteNumbering | undefined, no: number): string {
  if (opts?.customMark) return opts.customMark
  const start = opts?.numStart !== undefined && opts.numStart > 0 ? opts.numStart : 1
  return formatNoteMark(opts?.numFmt ?? 'decimal', no - 1 + start)
}

/** Word's built-in display when a document carries no note options */

/** document-wide note options of the loaded document (set before first render) */
let docNoteNumbering: { footnotes?: NoteNumbering; endnotes?: NoteNumbering } | undefined

export function setDocNoteNumbering(
  opts: { footnotes?: NoteNumbering; endnotes?: NoteNumbering } | undefined,
): void {
  docNoteNumbering = opts
}

/** options of one note kind (footnote → footnotes, endnote → endnotes) */
export function noteNumberingOfKind(
  kind: 'footnote' | 'endnote',
  source?: { footnotes?: NoteNumbering; endnotes?: NoteNumbering },
): NoteNumbering | undefined {
  const dicts = source ?? docNoteNumbering
  return kind === 'footnote' ? dicts?.footnotes : dicts?.endnotes
}

/** current marker of note #no under the loaded document's options */
export function docNoteMark(kind: 'footnote' | 'endnote', no: number): string {
  const opts = noteNumberingOfKind(kind)
  if (kind === 'footnote') return noteMarkerOf(opts, no)
  // endnotes default to lowercase roman when the file says nothing
  return noteMarkerOf(opts ?? { numFmt: 'lowerRoman' }, no)
}

/** fixed symbol of a kind under the loaded options, when one is set */
export function docNoteCustomMark(kind: 'footnote' | 'endnote'): string | null {
  return noteNumberingOfKind(kind)?.customMark ?? null
}
export function noteMarkText(kind: 'footnote' | 'endnote', no: number): string {
  return kind === 'endnote' ? toRoman(no) : String(no)
}

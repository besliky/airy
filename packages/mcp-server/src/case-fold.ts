// Index-safe case-insensitive literal matching, shared by the line sessions
// (markdown/html findReplace) and the docx ops engine (findReplace over runs,
// setMatchedFont occurrence styling). One primitive, three call sites: they
// must never drift apart again (BUG-1101 hit all of them at once).
//
// String.prototype.toLowerCase is NOT length-preserving: U+0130 (İ, Turkish
// dotted capital I) folds to "i" + U+0307, one code point becoming two code
// units. An index found in `text.toLowerCase()` therefore does not address
// `text` — slicing there silently deletes or duplicates characters after any
// İ earlier in the string. The fix: fold every code point on its own and
// remember, for each folded code unit, the index of the source code point
// that produced it, so match positions translate back into original
// coordinates before anything is sliced.
//
// Which matches are FOUND is unchanged: the folded subject is still exactly
// text.toLowerCase(). The per-code-point lengths align that string perfectly —
// İ is the only unconditional lowercase expansion, and the sole
// context-sensitive default-case rule (Greek Final_Sigma) replaces the
// character but stays 1:1 in length.
//
// A match that ends inside an expansion (find "i" against "İ") consumes the
// whole source code point: partial code points cannot be represented, and the
// alternative (slicing mid-character) would corrupt the string.

/** half-open [start, end) range of a literal match, in ORIGINAL coordinates */
export interface FoldRange {
  start: number
  end: number
}

/** code-unit width of the code point starting at `i` (surrogate pairs = 2) */
function codePointWidth(text: string, i: number): number {
  const code = text.charCodeAt(i)
  return code >= 0xd800 && code <= 0xdbff ? 2 : 1
}

/**
 * The folded text plus, for every folded code unit, the source index of the
 * code point that produced it. `origins` is null only if the per-code-point
 * walk and the whole-string fold disagree in length — impossible with the
 * Unicode default case conversion (see the module comment), but a mismatch
 * would corrupt every slice after it, so callers treat it as "no matches".
 */
function foldWithOrigins(text: string): { folded: string; origins: Int32Array | null } {
  const folded = text.toLowerCase()
  const origins = new Int32Array(folded.length)
  let out = 0
  let i = 0
  while (i < text.length) {
    const produced = String.fromCodePoint(text.codePointAt(i)!).toLowerCase().length
    for (let k = 0; k < produced; k++) origins[out + k] = i
    out += produced
    i += codePointWidth(text, i)
  }
  if (out !== folded.length) return { folded, origins: null }
  return { folded, origins }
}

/**
 * All non-overlapping case-insensitive literal matches of `find` in `text`,
 * as [start, end) ranges in ORIGINAL coordinates. An empty needle matches
 * nothing (callers validate non-empty finds, this is the backstop).
 */
export function caseInsensitiveRanges(text: string, find: string): FoldRange[] {
  const needle = find.toLowerCase()
  if (needle.length === 0) return []
  const { folded, origins } = foldWithOrigins(text)
  if (origins === null) return []
  const ranges: FoldRange[] = []
  let from = 0
  for (;;) {
    const at = folded.indexOf(needle, from)
    if (at === -1) break
    const last = at + needle.length - 1
    const start = origins[at]!
    ranges.push({ start, end: origins[last]! + codePointWidth(text, origins[last]!) })
    from = at + needle.length
  }
  return ranges
}

/**
 * Case-insensitive literal replace: the replacement text lands verbatim at
 * every match, everything between matches is copied from the original —
 * including İ and astral characters a lowered-index slice would have eaten.
 * Returns the rebuilt text plus the match count.
 */
export function replaceCaseInsensitive(
  text: string,
  find: string,
  replace: string,
): { text: string; count: number } {
  const ranges = caseInsensitiveRanges(text, find)
  if (ranges.length === 0) return { text, count: 0 }
  let out = ''
  let pos = 0
  for (const { start, end } of ranges) {
    out += text.slice(pos, start) + replace
    pos = end
  }
  return { text: out + text.slice(pos), count: ranges.length }
}

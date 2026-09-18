/**
 * Word-style "Use wildcards" find patterns, compiled to native RegExp, plus
 * the length-preserving diacritic folding behind "Ignore diacritics".
 *
 * Supported syntax (mirrors Word's Find and Replace wildcard mode):
 *   `?`          any single character (one Unicode code point)
 *   `*`          any run of characters, lazily — `s*d` finds `sad` inside
 *                `sad started` first, like Word's lazy wildcard matching
 *   `[a-m]`      one character in an inclusive code-point range
 *   `[!abc]`     one character not in the set; `[!a-m]` negates a range
 *   `\x`         escapes the next character (`\*` finds a literal `*`)
 * Inside `[ ]`, `!` is special only in leading position and `-` only
 * between two members (a leading or trailing `-` is literal); every other
 * character — including `?` and `*` — is literal.
 *
 * Unsupported Word operators (`( )` groups/backreferences, `{n}` counts,
 * `@`, `<`/`>` anchors) are treated as plain literals. A `[` without a
 * closing `]` falls back to a literal `[` instead of failing.
 *
 * `?`/`*` and classes match by code point, so CJK and emoji work; reversed
 * ranges like `[z-a]` simply match nothing (a negated reversed range matches
 * any character), matching Word's behavior.
 *
 * The block text fed to the matcher represents non-text inline nodes (hard
 * breaks, inline images/math/ruby/note refs) as a U+0000 placeholder. Word
 * never treats those objects as characters, so `?`, `*` and negated classes
 * all refuse to match the placeholder: only real text can be matched (and
 * thus only real text is ever inside a Replace range).
 */

/** regex-special characters that must be escaped when literal (u-flag safe) */
const RE_SPECIALS = /[.*+?^${}()|[\]\\]/g
/** inside a character class `-` is special too; `\-` is illegal outside one under the u flag */
const RE_CLASS_SPECIALS = /[.*+?^${}()|[\]\\-]/g
const esc = (ch: string) => ch.replace(RE_SPECIALS, '\\$&')
const escClass = (ch: string) => ch.replace(RE_CLASS_SPECIALS, '\\$&')

/** class that never matches (an empty `[]` is illegal under the u flag) */
const NEVER = '[^\\s\\S]'
/**
 * Class that matches any single code point except the U+0000 leaf placeholder
 * (see the header comment): `\u0000` is not a document character, so no
 * wildcard may match it (BUG-741). Valid under the `u` flag.
 */
const ANY = '[^\\u0000]'

interface ClassFragment {
  fragment: string | null
  /** pattern index just past the closing `]`; -1 when the class is unterminated */
  next: number
}

/** parse a `[...]` / `[!...]` class starting at `pcs[start] === '['` */
function compileClass(pcs: string[], start: number): ClassFragment {
  let k = start + 1
  const negate = pcs[k] === '!'
  if (negate) k++
  let inner = ''
  let atoms = 0
  while (k < pcs.length && pcs[k] !== ']') {
    let lo = pcs[k]
    if (lo === '\\' && k + 1 < pcs.length) {
      k++
      lo = pcs[k]
    }
    k++
    // `-` between two members forms a range unless it is last in the class
    if (pcs[k] === '-' && k + 1 < pcs.length && pcs[k + 1] !== ']') {
      let hi = pcs[k + 1]
      if (hi === '\\' && k + 2 < pcs.length) {
        k++
        hi = pcs[k + 1]
      }
      k += 2
      if (lo.codePointAt(0)! <= hi.codePointAt(0)!) {
        inner += `${escClass(lo)}-${escClass(hi)}`
        atoms++
      } // a reversed range contributes nothing, like Word
      continue
    }
    inner += escClass(lo)
    atoms++
  }
  if (k >= pcs.length) return { fragment: null, next: -1 }
  if (atoms === 0) return { fragment: negate ? ANY : NEVER, next: k + 1 }
  // a negated class also refuses the leaf placeholder, like `?` and `*`
  return { fragment: negate ? `[^${inner}\\u0000]` : `[${inner}]`, next: k + 1 }
}

/**
 * Compile a wildcard pattern to a RegExp with the `gu` flags (plus `i` when
 * `ignoreCase`). Returns null for an empty pattern. Zero-length matches
 * (e.g. a bare `*`) are the caller's concern: skip them and advance.
 */
export function compileWildcards(pattern: string, ignoreCase: boolean): RegExp | null {
  const pcs = Array.from(pattern)
  if (pcs.length === 0) return null
  let out = ''
  let i = 0
  while (i < pcs.length) {
    const c = pcs[i]
    if (c === '*') {
      out += `${ANY}*?`
      i++
    } else if (c === '?') {
      out += ANY
      i++
    } else if (c === '\\' && i + 1 < pcs.length) {
      out += esc(pcs[i + 1])
      i += 2
    } else if (c === '[') {
      const { fragment, next } = compileClass(pcs, i)
      if (fragment === null) {
        out += esc(c) // unterminated class: treat `[` as a literal
        i++
      } else {
        out += fragment
        i = next
      }
    } else {
      out += esc(c)
      i++
    }
  }
  try {
    return new RegExp(out, ignoreCase ? 'gui' : 'gu')
  } catch {
    return null
  }
}

const diacriticFoldCache = new Map<string, string>()

/**
 * Fold accented characters to their base form (NFD, combining marks
 * stripped). Strictly length-preserving per code point — every source code
 * point maps to exactly one code point of the same UTF-16 length — so match
 * offsets computed on folded text address the original text directly. A lone
 * combining mark (no base within the same code point) and decompositions
 * that do not leave exactly one base character (e.g. precomposed Hangul)
 * are kept as-is. Case is untouched; pair with case folding / the regex `i`
 * flag for case-insensitive search.
 */
export function foldDiacritics(s: string): string {
  let out = ''
  for (const ch of s) {
    let fold = diacriticFoldCache.get(ch)
    if (fold === undefined) {
      fold = ch
      const nfd = ch.normalize('NFD')
      if (nfd !== ch) {
        let base = ''
        let count = 0
        for (const p of nfd) {
          if (!/\p{M}/u.test(p)) {
            base += p
            count++
          }
        }
        if (count === 1 && base.length === ch.length) fold = base
      }
      diacriticFoldCache.set(ch, fold)
    }
    out += fold
  }
  return out
}

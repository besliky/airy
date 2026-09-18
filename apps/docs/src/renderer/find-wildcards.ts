/**
 * Word-style "Use wildcards" find patterns, compiled to native RegExp.
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
 */

/** regex-special characters that must be escaped when literal (u-flag safe) */
const RE_SPECIALS = /[.*+?^${}()|[\]\\]/g
/** inside a character class `-` is special too; `\-` is illegal outside one under the u flag */
const RE_CLASS_SPECIALS = /[.*+?^${}()|[\]\\-]/g
const esc = (ch: string) => ch.replace(RE_SPECIALS, '\\$&')
const escClass = (ch: string) => ch.replace(RE_CLASS_SPECIALS, '\\$&')

/** class that never matches (an empty `[]` is illegal under the u flag) */
const NEVER = '[^\\s\\S]'
/** class that matches any single code point */
const ANY = '[\\s\\S]'

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
  return { fragment: negate ? `[^${inner}]` : `[${inner}]`, next: k + 1 }
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

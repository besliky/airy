/**
 * Word-style "Use wildcards" find patterns, compiled to a small token program
 * and executed by a linear matcher, plus the length-preserving folding behind
 * "Ignore diacritics" and case-insensitive search.
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
 *
 * Execution is a right-to-left dynamic program over the code points of the
 * searched text (O(tokens x text)), not a backtracking regexp: patterns made
 * of several `*` runs used to compile to chained lazy `[\s\S]*?` quantifiers
 * whose overlapping backtracking froze the renderer on large documents
 * (`*a*a*a*a*z` over 50 KB never finished — BUG-742). The program reproduces
 * the regexp engine's preference order exactly: leftmost match, and each `*`
 * stops at the first position where the rest of the pattern matches.
 */

/**
 * The U+0000 stand-in for a non-text inline node in the flattened block
 * text (see the header comment): no wildcard construct ever matches it.
 */
const PLACEHOLDER = 0

/** one `[a-m]` / `[!abc]` member set: inclusive code-point ranges */
type ClassRange = [number, number]

type WildToken =
  | { kind: 'lit'; cps: number[] } // literal run (pre-folded when ignoreCase)
  | { kind: 'any' } // `?`
  | { kind: 'star' } // `*`
  | { kind: 'class'; negated: boolean; ranges: ClassRange[] }

interface ClassProgram {
  /** null when the class is unterminated (the `[` stays a literal) */
  ranges: ClassRange[] | null
  negate: boolean
  /** pattern index just past the closing `]`; -1 when the class is unterminated */
  next: number
}

/** parse a `[...]` / `[!...]` class starting at `pcs[start] === '['` */
function parseClass(pcs: string[], start: number): ClassProgram {
  let k = start + 1
  const negate = pcs[k] === '!'
  if (negate) k++
  const ranges: ClassRange[] = []
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
      const loCp = lo.codePointAt(0)!
      const hiCp = hi.codePointAt(0)!
      // a reversed range contributes nothing, like Word
      if (loCp <= hiCp) ranges.push([loCp, hiCp])
      continue
    }
    const cp = lo.codePointAt(0)!
    ranges.push([cp, cp])
  }
  if (k >= pcs.length) return { ranges: null, negate, next: -1 }
  return { ranges, negate, next: k + 1 }
}

/** does a class token match one (already folded) code point? */
function classMatches(token: { negated: boolean; ranges: ClassRange[] }, cp: number): boolean {
  if (cp === PLACEHOLDER) return false
  let inSet = false
  for (const [lo, hi] of token.ranges) {
    if (cp >= lo && cp <= hi) {
      inSet = true
      break
    }
  }
  return token.negated ? !inSet : inSet
}

function parseTokens(pattern: string, ignoreCase: boolean): WildToken[] {
  const pcs = Array.from(pattern)
  const tokens: WildToken[] = []
  const fold = (cp: number) => (ignoreCase ? foldCaseCp(cp) : cp)
  let lit: number[] | null = null
  const flush = () => {
    if (lit) {
      tokens.push({ kind: 'lit', cps: lit })
      lit = null
    }
  }
  let i = 0
  while (i < pcs.length) {
    const c = pcs[i]
    if (c === '*') {
      flush()
      tokens.push({ kind: 'star' })
      i++
    } else if (c === '?') {
      flush()
      tokens.push({ kind: 'any' })
      i++
    } else if (c === '\\' && i + 1 < pcs.length) {
      ;(lit ??= []).push(fold(pcs[i + 1].codePointAt(0)!))
      i += 2
    } else if (c === '[') {
      const { ranges, negate, next } = parseClass(pcs, i)
      if (ranges === null) {
        ;(lit ??= []).push(fold(c.codePointAt(0)!)) // unterminated class: literal `[`
        i++
      } else {
        flush()
        tokens.push({ kind: 'class', negated: negate, ranges })
        i = next
      }
    } else {
      ;(lit ??= []).push(fold(c.codePointAt(0)!))
      i++
    }
  }
  flush()
  return tokens
}

/** one non-overlapping match, in UTF-16 offsets (`end` exclusive) */
export interface WildcardMatch {
  start: number
  end: number
}

export interface WildcardProgram {
  /**
   * All non-empty, non-overlapping matches in `hay`, leftmost first, with
   * each lazy `*` stopping at the first position where the rest matches —
   * the same order a `gu` regexp used to produce. Zero-length matches
   * (e.g. a bare `*`) are skipped, like the old exec loop did.
   */
  execAll(hay: string): WildcardMatch[]
}

/**
 * Compile a wildcard pattern to a match program. Returns null for an empty
 * pattern. `ignoreCase` folds the pattern up front and the haystack per code
 * point with the same length-preserving case fold, so offsets stay exact.
 */
export function compileWildcards(pattern: string, ignoreCase: boolean): WildcardProgram | null {
  const tokens = parseTokens(pattern, ignoreCase)
  if (tokens.length === 0) return null
  return {
    execAll: (hay: string) => execTokens(tokens, hay, ignoreCase),
  }
}

/**
 * The Word wildcard operators this engine treats as plain literals, for the
 * find panel's inline warning: `( )` groups/backreferences, `{ }` repeat
 * counts, `@` one-or-more and `<`/`>` word anchors (the header list).
 * Escaped occurrences (`\(`) and everything inside `[...]` (where those
 * characters are literal by Word's rules) are not reported. `+` is NOT a
 * Word wildcard operator — it is a literal in Word too — so it is not
 * flagged: warning on every literal plus would drown legit searches
 * ("C++") while Word itself stays silent there.
 *
 * Returns the distinct operator characters in first-occurrence order; empty
 * when the pattern uses only the supported subset. Pure and linear — safe to
 * run per keystroke next to the debounced scan.
 */
export function unsupportedWildcardOperators(pattern: string): string[] {
  const found: string[] = []
  const pcs = Array.from(pattern)
  for (let i = 0; i < pcs.length; i++) {
    const c = pcs[i]!
    if (c === '\\' && i + 1 < pcs.length) {
      i++ // escaped: the next character is a literal
    } else if (c === '[') {
      // inside a class every other operator is literal; an unterminated `[`
      // stays a literal `[` and the scan resumes after it (like parseTokens)
      let k = i + 1
      if (pcs[k] === '!') k++
      if (pcs[k] === '\\') k++
      while (k < pcs.length && pcs[k] !== ']') {
        if (pcs[k] === '\\' && k + 1 < pcs.length) k++
        k++
      }
      i = k < pcs.length ? k : i // terminated: jump past `]`; else just past `[`
    } else if ('(){}@<>'.includes(c)) {
      if (!found.includes(c)) found.push(c)
    }
  }
  return found
}

/**
 * The linear matcher. For every token position `ti` (right to left) it
 * computes `f[s]` = end slot of the match of `tokens[ti..]` starting at
 * code-point slot `s` (or -1), from the already computed layer of `ti + 1`:
 * one pass per token, O(tokens x text) overall — no backtracking search, so
 * hostile patterns cannot blow up (BUG-742).
 */
function execTokens(tokens: WildToken[], hay: string, ignoreCase: boolean): WildcardMatch[] {
  const n = hay.length
  const matches: WildcardMatch[] = []
  if (n === 0) return matches
  // code-point slots with their UTF-16 start offsets, folded for comparison
  const cps: number[] = []
  const starts: number[] = []
  for (let i = 0; i < n;) {
    const cp = hay.codePointAt(i)!
    cps.push(ignoreCase ? foldCaseCp(cp) : cp)
    starts.push(i)
    i += cp > 0xffff ? 2 : 1
  }
  const m = cps.length
  // g holds the computed layer of token index ti + 1; f is the scratch buffer
  // the current pass writes into (they swap at the end of each pass)
  let f = new Int32Array(m + 1)
  let g = new Int32Array(m + 1)
  for (let s = 0; s <= m; s++) g[s] = s // empty suffix matches empty everywhere
  for (let ti = tokens.length - 1; ti >= 0; ti--) {
    const token = tokens[ti]
    f.fill(-1)
    if (token.kind === 'star') {
      // lazy: the run stops at the first slot >= s where the rest matches;
      // it may not cross a placeholder (and never consumes one)
      let nextGood = -1 // first slot >= s with g != -1
      let firstNul = m // first placeholder slot >= s (m = none)
      for (let s = m; s >= 0; s--) {
        if (cps[s] === PLACEHOLDER) firstNul = s
        if (g[s] !== -1) nextGood = s
        f[s] = nextGood !== -1 && nextGood <= firstNul ? g[nextGood] : -1
      }
    } else if (token.kind === 'any') {
      for (let s = 0; s < m; s++) {
        if (cps[s] !== PLACEHOLDER) f[s] = g[s + 1]
      }
    } else if (token.kind === 'class') {
      for (let s = 0; s < m; s++) {
        if (classMatches(token, cps[s])) f[s] = g[s + 1]
      }
    } else {
      const run = token.cps
      const k = run.length
      for (let s = 0; s + k <= m; s++) {
        let ok = true
        for (let j = 0; j < k; j++) {
          const cp = cps[s + j]
          if (cp !== run[j] || cp === PLACEHOLDER) {
            ok = false
            break
          }
        }
        if (ok) f[s] = g[s + k]
      }
    }
    const swap = f
    f = g
    g = swap // g is now the freshly computed layer of token index ti
  }
  // collect: from each slot the first position with a non-empty match wins,
  // then the scan continues past its end (the old exec-loop's lastIndex rule)
  for (let s = 0; s < m;) {
    const end = g[s]
    if (end > s) {
      matches.push({ start: starts[s], end: end < m ? starts[end] : n })
      s = end
    } else {
      s++ // no match at s (or only a zero-length one, e.g. a bare `*`)
    }
  }
  return matches
}

const caseFoldCache = new Map<number, number>()

/**
 * Length-preserving lowercase for one code point: chars whose lowercase grows
 * (`İ` → `i̇`) stay as-is so UTF-16 offsets never shift.
 */
function foldCaseCp(cp: number): number {
  let fold = caseFoldCache.get(cp)
  if (fold === undefined) {
    const ch = String.fromCodePoint(cp)
    const lower = ch.toLowerCase()
    fold = lower.length === ch.length ? lower.codePointAt(0)! : cp
    caseFoldCache.set(cp, fold)
  }
  return fold
}

/** length-preserving lowercase: chars whose lowercase grows ('İ' → 'i̇') stay as-is so match offsets never shift */
export function foldCase(s: string): string {
  let out = ''
  for (const ch of s) {
    const lower = ch.toLowerCase()
    out += lower.length === ch.length ? lower : ch
  }
  return out
}

const diacriticFoldCache = new Map<string, string>()

/**
 * Fold accented characters to their base form (NFD, combining marks
 * stripped). Strictly length-preserving per code point — every source code
 * point maps to exactly one code point of the same UTF-16 length — so match
 * offsets computed on folded text address the original text directly. A lone
 * combining mark (no base within the same code point) and decompositions
 * that do not leave exactly one base character (e.g. precomposed Hangul)
 * are kept as-is. Case is untouched; pair with case folding / the matcher's
 * `ignoreCase` for case-insensitive search.
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

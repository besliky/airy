/**
 * Word-parity text statistics.
 *
 * Word's CJK rule: Words = Asian characters (counted one by one, punctuation
 * included) + non-Asian words (whitespace/punct-delimited). The dialog also
 * reports the two addends separately — Chinese users care about the
 * Asian-character figure.
 */

import type { Node as PmNode } from '@tiptap/pm/model'

// Han (incl. radicals/compat/ext-B+), kana, hangul, bopomofo, CJK symbols
// and punctuation (U+3001 up: the ideographic space stays whitespace),
// fullwidth forms
const ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟、-〿぀-ヿ㄀-ㄯ㄰-㆏㇀-ㇿ㐀-䶿一-鿿가-힯豈-﫿！-｠￠-￦]|[\uD840-\uD87F][\uDC00-\uDFFF]/g

const NON_ASIAN_WORD_RE = /[A-Za-z0-9À-ɏ]+(?:['-][A-Za-z0-9À-ɏ]+)*/g

export function asianCharCount(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length
}

export function nonAsianWordCount(text: string): number {
  return (text.match(NON_ASIAN_WORD_RE) ?? []).length
}

/** Word's Words figure: asian chars + non-asian words */
/** per-block word counts: a block's words never cross its boundary, so unchanged blocks reuse their cached count */
const blockWordCache = new WeakMap<PmNode, number>()

/**
 * Whole-document word count with per-block memoization. Equivalent to
 * `countWords(doc.textContent)` per block and summed; unlike a naive
 * concatenation it never merges the last word of a block with the first word
 * of the next one (Word counts them separately too). PERF-1639: the
 * per-document computation rebuilt the full text string on every transaction,
 * which was O(document) per streamed chunk on large files.
 */
export function countWordsInDoc(doc: PmNode): number {
  let total = 0
  doc.forEach((block) => {
    let n = blockWordCache.get(block)
    if (n === undefined) {
      n = countWords(block.textContent)
      blockWordCache.set(block, n)
    }
    total += n
  })
  return total
}

export function countWords(text: string): number {
  return asianCharCount(text) + nonAsianWordCount(text)
}

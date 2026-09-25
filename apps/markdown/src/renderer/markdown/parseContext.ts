/**
 * Parse/serialize context fixes for @tiptap/markdown (BUG-1703). The manager
 * exposes no pre/post hooks, so these two pure transforms are wrapped around
 * `editor.markdown.parse` / `editor.markdown.serialize` (see extensions.ts).
 *
 * 1. `liftIndentedCodeAfterLists` — marked folds a blank-line-separated block
 *    indented past a list's content column into the last list item as a plain
 *    paragraph. In that position the block is authored code (BUG-1742's
 *    fixture: a "plain four-space code block" after an ordered list), and the
 *    model must not swallow it: list edits and the docx export would treat
 *    the code as item text, and the serializer re-indents it (+2 spaces on
 *    the audit vector). The transform rewrites only that region into a fenced
 *    block before parsing, so the document model gains a real code block.
 *    The fence materializes in the file on the next save — the manager has no
 *    indented-code serializer, so the fenced form is the canonical on-disk
 *    shape for a code block in this position.
 *
 *    The rule works in two indent tiers, and is deliberately narrow to keep
 *    ordinary loose-list continuation paragraphs untouched:
 *
 *    - Two units deep (8 spaces of the configured 4-space indentation) the
 *      block always lifts: no serializer continuation lives that deep.
 *    - One unit deep (4 spaces) the chunk lifts only when it cannot be a
 *      serializer continuation: it needs at least two content lines (a single
 *      line is exactly the serializer's nested-paragraph shape
 *      "1. one\n\n    more about one") and no line may carry the serializer's
 *      hard-break form (text followed by 2+ trailing spaces — "serialize"
 *      emits wrapped list-item paragraphs as "    line one  \n    line two").
 *
 *    In both tiers the block must follow a blank line at the tail of the list
 *    context. Consecutive indented chunks separated by blank lines are one
 *    code block (indented-code chunk semantics).
 *
 * 2. `stripBlankLinePadding` — an empty paragraph nested in a list item
 *    serializes as a line of indentation spaces ("    "), i.e. trailing
 *    whitespace in the saved file. Whitespace-only lines outside literal
 *    regions (fenced code, raw HTML containers, HTML comments, block math)
 *    are blanked. Lines with text are never touched, so a hard break's
 *    trailing double-space survives.
 */

const FENCE = /^ {0,3}(`{3,}|~{3,})/
const HTML_RAW_OPEN = /^ {0,3}<(script|pre|style|textarea)\b/i
const HTML_RAW_CLOSE = /<\/(?:script|pre|style|textarea)>/i
const COMMENT_OPEN = /^ {0,3}<!--/
const COMMENT_CLOSE = /-->/
/** list item marker at column <= 3 (deeper indents are item content) */
const LIST_MARKER = /^ {0,3}(?:[-+*]|\d{1,9}[.)])(?:\s|$)/
/** block math delimiter ($$ ... $$), serialized on its own lines */
const BLOCK_MATH = /^\$\$/

/** the lifted block's two-unit form: indented by two 4-space nesting units */
const CODE_INDENT = 8

/** the lifted block's one-unit form (BUG-1742 audit vector): 4 spaces */
const ONE_UNIT_INDENT = 4

/** the serializer's hard-break form: text followed by 2+ trailing spaces */
const SERIALIZER_HARD_BREAK = / {2,}$/

const indentOf = (line: string): number => line.length - line.trimStart().length
const isBlank = (line: string): boolean => line.trim() === ''

interface ScannerState {
  fence: { char: string; length: number } | null
  htmlTag: string | null
  comment: boolean
}

/** advance the literal-region state over one line; true while inside one */
function inLiteralRegion(state: ScannerState, line: string): boolean {
  if (state.fence) {
    const close = FENCE.exec(line)
    if (close && close[1][0] === state.fence.char && close[1].length >= state.fence.length) {
      state.fence = null
    }
    return true
  }
  if (state.htmlTag) {
    if (HTML_RAW_CLOSE.test(line)) state.htmlTag = null
    return true
  }
  if (state.comment) {
    if (COMMENT_CLOSE.test(line)) state.comment = false
    return true
  }
  return false
}

function enterLiteralRegion(state: ScannerState, line: string): boolean {
  const openFence = FENCE.exec(line)
  if (openFence) {
    state.fence = { char: openFence[1][0], length: openFence[1].length }
    return true
  }
  const openHtml = HTML_RAW_OPEN.exec(line)
  if (openHtml) {
    if (!HTML_RAW_CLOSE.test(line)) state.htmlTag = openHtml[1]
    return true
  }
  if (COMMENT_OPEN.test(line)) {
    if (!COMMENT_CLOSE.test(line)) state.comment = true
    return true
  }
  return false
}

/** longest run of backticks anywhere in the text (fence-collision guard) */
function maxBacktickRun(text: string): number {
  let max = 0
  for (const run of text.match(/`+/g) ?? []) max = Math.max(max, run.length)
  return max
}

/**
 * Should the indented chunk lines[from..to] lift out of the list as code?
 * Two-unit chunks always do (no serializer continuation lives that deep —
 * though a serializer hard-break trail still vetoes, it marks wrapped text).
 * One-unit chunks lift only when they cannot be a serializer continuation:
 * at least two content lines and none carrying a hard-break trail.
 */
function liftsAsCode(lines: string[], from: number, to: number, twoUnit: boolean): boolean {
  let contentLines = 0
  for (let j = from; j <= to; j++) {
    if (isBlank(lines[j])) continue
    contentLines++
    if (SERIALIZER_HARD_BREAK.test(lines[j])) return false
  }
  return twoUnit || contentLines >= 2
}

export function liftIndentedCodeAfterLists(markdown: string): string {
  const lines = markdown.split('\n')
  const out: string[] = []
  const state: ScannerState = { fence: null, htmlTag: null, comment: false }
  let inList = false

  let i = 0
  while (i < lines.length) {
    const line = lines[i]

    if (inLiteralRegion(state, line) || enterLiteralRegion(state, line)) {
      out.push(line)
      i++
      continue
    }
    if (isBlank(line)) {
      out.push(line)
      i++
      continue
    }

    const indent = indentOf(line)
    const atColumnZero = indent < ONE_UNIT_INDENT
    const isMarker = atColumnZero && LIST_MARKER.test(line)

    // candidate: indented block, list still open, blank line before it
    if (indent >= ONE_UNIT_INDENT && inList && isBlank(out[out.length - 1] ?? '')) {
      const twoUnit = indent >= CODE_INDENT
      // chunk membership threshold: lines indented by at least one unit below
      // the chunk's tier (a two-unit chunk is not joined by one-unit lines)
      const chunkIndent = twoUnit ? CODE_INDENT : ONE_UNIT_INDENT
      // collect chunks: indented lines, blank lines between them kept
      // when another indented line follows (indented-code chunk form)
      let lastContent = -1
      for (let j = i; j < lines.length; j++) {
        if (isBlank(lines[j])) continue
        if (indentOf(lines[j]) < chunkIndent) break
        lastContent = j
      }
      if (lastContent >= i && liftsAsCode(lines, i, lastContent, twoUnit)) {
        const content = lines
          .slice(i, lastContent + 1)
          .map((l) => (isBlank(l) ? '' : l.slice(chunkIndent)))
        // a closing fence is any line starting with a run at least as long as
        // the fence itself — pick a length no content line can match
        const fenceText = '`'.repeat(Math.max(3, maxBacktickRun(content.join('\n')) + 1))
        out.push(fenceText, ...content, fenceText)
        i = lastContent + 1
        // the fence is a top-level block: the list context is gone
        inList = false
        continue
      }
    }

    if (inList && atColumnZero && !isMarker) inList = false
    if (isMarker) inList = true
    out.push(line)
    i++
  }
  return out.join('\n')
}

export function stripBlankLinePadding(markdown: string): string {
  const lines = markdown.split('\n')
  const state: ScannerState = { fence: null, htmlTag: null, comment: false }
  let math = false

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]
    if (inLiteralRegion(state, line)) continue
    if (math) {
      if (BLOCK_MATH.test(line.trim())) math = false
      continue
    }
    if (enterLiteralRegion(state, line)) continue
    if (BLOCK_MATH.test(line.trim())) {
      math = true
      continue
    }
    // whitespace-only line in plain context = an indented empty paragraph;
    // blank it so the file carries no trailing whitespace
    if (isBlank(line)) lines[i] = ''
  }
  return lines.join('\n')
}

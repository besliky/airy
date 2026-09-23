/**
 * Safe top-level segmentation of a markdown body for phased hydration
 * (PERF-1647). `@tiptap/markdown`'s parser is quadratic in the size of its
 * input — a 20k-paragraph file spends tens of seconds inside one
 * `markdown.parse` call — while the ProseMirror document build itself is
 * sub-second. Parsing the body chunk by chunk at true top-level block
 * boundaries keeps the parse cost linear.
 *
 * The scanner only ever cuts BETWEEN top-level blocks, and only where a run
 * of blank lines separates blocks and no construct is left open that a later
 * segment would need:
 * - fenced code blocks (``` / ~~~, indent <= 3),
 * - raw HTML blocks (<pre>/<script>/<style>/<textarea>) and HTML comments,
 * - lists (a blank line between items keeps the list open — a loose list must
 *   not be cut or the serialized file would gain a list boundary),
 * - link reference definitions (resolved document-wide by the parser, so a
 *   body containing one is never segmented at all).
 *
 * Segments are string slices of the body between line offsets, so
 * concatenating all segments reproduces the input byte for byte.
 */

export interface SegmentOptions {
  /** maximum lines accumulated before a boundary is taken (when one exists) */
  maxLines?: number
  /** maximum characters accumulated before a boundary is taken (when one exists) */
  maxChars?: number
}

const DEFAULT_MAX_LINES = 2000
const DEFAULT_MAX_CHARS = 512 * 1024

const FENCE = /^ {0,3}(`{3,}|~{3,})/
const HTML_RAW_OPEN = /^ {0,3}<(script|pre|style|textarea)\b/i
const HTML_RAW_CLOSE = /<\/(?:script|pre|style|textarea)>/i
const COMMENT_OPEN = /^ {0,3}<!--/
const COMMENT_CLOSE = /-->/
/** link reference definition — resolved across the whole document by the parser */
const REF_DEF = /^ {0,3}\[[^\]]+]:\s*\S/
/** list item marker at column <= 3 (indented text is continuation, not a marker) */
const LIST_MARKER = /^ {0,3}(?:[-+*]|\d{1,9}[.)])(?:\s|$)/

export function splitBodyForHydration(body: string, options: SegmentOptions = {}): string[] {
  const maxLines = options.maxLines ?? DEFAULT_MAX_LINES
  const maxChars = options.maxChars ?? DEFAULT_MAX_CHARS
  const lines = body.split('\n')

  // reference definitions are resolved document-wide by marked — a body that
  // uses them must be parsed as one piece or references would break
  if (lines.some((line) => REF_DEF.test(line))) return [body]

  // char offset at which each line starts; starts[lines.length] = end of body
  const starts: number[] = [0]
  for (const line of lines) starts.push(starts[starts.length - 1] + line.length + 1)

  const segments: string[] = []
  let startLine = 0
  let fence: { char: string; length: number } | null = null
  let htmlTag: string | null = null
  let comment = false
  /** inside a list item run (item, indented continuation, or blank + item) */
  let inList = false
  /** the previous line(s) were blank — a boundary may exist before a real block */
  let blankRun = false

  const flush = (lineIndex: number): void => {
    segments.push(body.slice(starts[startLine], starts[lineIndex]))
    startLine = lineIndex
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]

    if (fence) {
      const close = FENCE.exec(line)
      if (close && close[1][0] === fence.char && close[1].length >= fence.length) fence = null
      continue
    }
    if (htmlTag) {
      if (HTML_RAW_CLOSE.test(line)) htmlTag = null
      continue
    }
    if (comment) {
      if (COMMENT_CLOSE.test(line)) comment = false
      continue
    }

    const openFence = FENCE.exec(line)
    if (openFence) {
      fence = { char: openFence[1][0], length: openFence[1].length }
      blankRun = false
      continue
    }
    const openHtml = HTML_RAW_OPEN.exec(line)
    if (openHtml) {
      if (!HTML_RAW_CLOSE.test(line)) htmlTag = openHtml[1]
      blankRun = false
      continue
    }
    if (COMMENT_OPEN.test(line)) {
      if (!COMMENT_CLOSE.test(line)) comment = true
      blankRun = false
      continue
    }

    if (line.trim() === '') {
      blankRun = true
      continue
    }

    // first non-blank line after blanks: it decides whether an open list
    // survives the blank line, and whether a boundary may be taken before it
    const atColumnZero = /^[^\s]/.test(line)
    const isMarker = atColumnZero && LIST_MARKER.test(line)
    if (blankRun && inList && atColumnZero && !isMarker) {
      // a blank line followed by a column-0 non-marker line closes the list;
      // a marker or an indented continuation keeps it open
      inList = false
    }
    // boundary: blanks separate blocks, nothing open, list closed, and this
    // line starts a fresh top-level block (column 0, not a list item)
    const canCut =
      blankRun && i > startLine && lines[i - 1] === '' && !inList && atColumnZero && !isMarker
    const overBudget = i - startLine >= maxLines || starts[i] - starts[startLine] >= maxChars
    if (canCut && overBudget) {
      // cut at the START of the trailing blank run: the separator blanks lead
      // the next segment (leading blanks never parse into nodes, while
      // trailing blanks after a list would parse into an extra empty
      // paragraph and break parity with the monolithic parse)
      let cut = i - 1
      while (cut > startLine && lines[cut - 1] === '') cut--
      flush(cut)
      inList = false
    }
    if (isMarker) inList = true
    blankRun = false
  }
  segments.push(body.slice(starts[startLine]))
  return segments
}

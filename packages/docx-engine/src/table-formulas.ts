// Table formula fields (=SUM(ABOVE) …), Word's Table Layout → Formula dialog.
//
// Word reference semantics implemented here:
// - Storage: <w:fldSimple w:instr=" =SUM(ABOVE) \# &quot;#,##0.00&quot; "> whose
//   child runs are the cached result shown until the field updates (F9).
// - Direction operands (ABOVE/BELOW/LEFT/RIGHT) select the contiguous cells
//   nearest the formula cell while they hold numbers; a blank/text cell stops
//   the scan, and repeating-header rows (w:tblHeader) are ignored.
// - Explicit cell references (A1, B2:C4) skip blank/text cells instead of
//   stopping (Word sums the numeric cells of the range).
// - A direction that finds no numeric cells (e.g. an empty column) produces
//   Word's literal "Undefined Bookmark" error text.
// - Numeric picture switch \# "…" with the minimal Word code set: 0 # x
//   placeholders, . decimal, , grouping, % (value × 100), quoted/escaped
//   literals and up to three positive;negative;zero sections.

/** Word's field result when a positional operand selects no cells at all */
export const FORMULA_EMPTY_ERROR = 'Undefined Bookmark'
/** Word's field result for a malformed formula */
export const FORMULA_SYNTAX_ERROR = '!Syntax Error'
/** Word's field result when a formula divides by zero */
export const FORMULA_DIV_ZERO_ERROR = '!Division By Zero'

export type FormulaDirection = 'ABOVE' | 'BELOW' | 'LEFT' | 'RIGHT'

/** cell texts of the physical table grid: texts[row][col] */
export type FormulaGridTexts = readonly (readonly string[])[]
/** header-row predicate (rows marked w:tblHeader are ignored by direction scans) */
export type HeaderRowPredicate = (row: number) => boolean

export interface FormulaGrid {
  texts: FormulaGridTexts
  isHeaderRow?: HeaderRowPredicate
}

interface ParsedFormula {
  expr: string
  picture: string | null
}

/**
 * Split a formula field instruction into the expression body (after '=') and
 * the numeric picture of its \# switch. Returns null when the instruction is
 * not a formula (no leading '='). Word writes the switch as
 * `\# "picture"` (double quotes), but single quotes and bare words occur.
 */
export function parseFormulaInstruction(instr: string): ParsedFormula | null {
  const trimmed = instr.trim()
  if (!trimmed.startsWith('=')) return null
  const expr = trimmed.slice(1)
  // the picture runs from \# to the next switch or the end; a quoted picture
  // keeps spaces and escaped quotes
  const pic = /\s\\#\s*(?:"((?:[^"\\]|\\.)*)"|'([^']*)'|([^\\\s]+))/.exec(expr)
  let picture: string | null = null
  let body = expr
  if (pic) {
    picture = pic[1] !== undefined ? pic[1].replace(/\\"/g, '"') : (pic[2] ?? pic[3] ?? '')
    body = expr.slice(0, pic.index)
  }
  // any other switch (\* MERGEFORMAT, \! ...) ends the expression like Word
  const switchAt = /(?:^|\s)\\[A-Za-z*]/.exec(body)
  if (switchAt) body = body.slice(0, switchAt.index)
  // trailing/leading whitespace inside the instruction is Word's own spelling
  return { expr: body.trim(), picture }
}

/** numeric value of a table cell's text, or null when the cell is not numeric */
export function parseCellNumber(text: string): number | null {
  const s = text.replace(/[\s\u00a0\u2007\u202f]/g, '')
  if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(s)) return null
  return parseFloat(s)
}

/** Direction keyword as accepted inside function arguments (case-insensitive) */
function asDirection(name: string): FormulaDirection | null {
  const upper = name.toUpperCase()
  return upper === 'ABOVE' || upper === 'BELOW' || upper === 'LEFT' || upper === 'RIGHT'
    ? (upper as FormulaDirection)
    : null
}

/**
 * Word's positional operand: the contiguous run of cells in `dir` starting at
 * the neighbor of (row, col), kept while the cells hold numbers. Blank or
 * text cells stop the scan; repeating-header rows are skipped (Word ignores
 * heading rows for positional arguments).
 */
export function collectDirectionOperands(
  texts: FormulaGridTexts,
  row: number,
  col: number,
  dir: FormulaDirection,
  isHeaderRow?: HeaderRowPredicate,
): number[] {
  const out: number[] = []
  if (dir === 'LEFT' || dir === 'RIGHT') {
    const dc = dir === 'LEFT' ? -1 : 1
    const rowTexts = texts[row] ?? []
    for (let c = col + dc; c >= 0 && c < rowTexts.length; c += dc) {
      const n = parseCellNumber(rowTexts[c] ?? '')
      if (n === null) break
      out.push(n)
    }
  } else {
    const dr = dir === 'ABOVE' ? -1 : 1
    for (let r = row + dr; r >= 0 && r < texts.length; r += dr) {
      if (isHeaderRow?.(r)) continue
      const n = parseCellNumber(texts[r]?.[col] ?? '')
      if (n === null) break
      out.push(n)
    }
  }
  return out
}

/** A1-style cell reference → [row, col], or null */
function parseCellRef(ref: string): [number, number] | null {
  const m = /^([A-Za-z]{1,3})([0-9]+)$/.exec(ref)
  if (!m) return null
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return [parseInt(m[2], 10) - 1, col - 1]
}

// ---- expression evaluation ----

class FormulaEvalError extends Error {
  constructor(readonly kind: 'empty' | 'div0' | 'syntax') {
    super(kind)
  }
}

type EvalValue = number | number[]

interface Token {
  kind: 'num' | 'ident' | 'op'
  text: string
}

function tokenize(src: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (/\s/.test(ch)) {
      i++
      continue
    }
    if (/[0-9.]/.test(ch)) {
      const m = /^[0-9]*\.?[0-9]+/.exec(src.slice(i))
      if (!m) throw new FormulaEvalError('syntax')
      tokens.push({ kind: 'num', text: m[0] })
      i += m[0].length
      continue
    }
    if (/[A-Za-z_]/.test(ch)) {
      const m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(src.slice(i))!
      tokens.push({ kind: 'ident', text: m[0] })
      i += m[0].length
      continue
    }
    const two = src.slice(i, i + 2)
    if (two === '<=' || two === '>=' || two === '<>') {
      tokens.push({ kind: 'op', text: two })
      i += 2
      continue
    }
    if ('+-*/^(),:<>=%'.includes(ch)) {
      tokens.push({ kind: 'op', text: ch })
      i++
      continue
    }
    throw new FormulaEvalError('syntax')
  }
  return tokens
}

/** grammar precedence levels, loosest first */
function tokenizeAndParse(src: string): (scope: EvalScope) => EvalValue {
  const tokens = tokenize(src)
  let pos = 0
  const peek = () => tokens[pos]
  const eat = (text?: string): Token => {
    const t = tokens[pos]
    if (!t || (text !== undefined && t.text !== text)) throw new FormulaEvalError('syntax')
    pos++
    return t
  }
  const atOp = (...texts: string[]) => peek()?.kind === 'op' && texts.includes(peek()!.text)

  const parseExpr = (): ((scope: EvalScope) => EvalValue) => {
    let left = parseSum()
    while (atOp('=', '<', '>', '<=', '>=', '<>')) {
      const op = eat().text
      const right = parseSum()
      const prev = left
      left = (scope) => {
        const a = toNumber(prev(scope))
        const b = toNumber(right(scope))
        const truth =
          op === '='
            ? a === b
            : op === '<>'
              ? a !== b
              : op === '<'
                ? a < b
                : op === '<='
                  ? a <= b
                  : op === '>'
                    ? a > b
                    : a >= b
        return truth ? 1 : 0
      }
    }
    return left
  }
  const parseSum = (): ((scope: EvalScope) => EvalValue) => {
    let left = parseTerm()
    while (atOp('+', '-')) {
      const op = eat().text
      const right = parseTerm()
      const prev = left
      left = (scope) => {
        const a = toNumber(prev(scope))
        const b = toNumber(right(scope))
        return op === '+' ? a + b : a - b
      }
    }
    return left
  }
  const parseTerm = (): ((scope: EvalScope) => EvalValue) => {
    let left = parseUnary()
    while (atOp('*', '/')) {
      const op = eat().text
      const right = parseUnary()
      const prev = left
      left = (scope) => {
        const a = toNumber(prev(scope))
        const b = toNumber(right(scope))
        if (op === '*') return a * b
        if (b === 0) throw new FormulaEvalError('div0')
        return a / b
      }
    }
    return left as (scope: EvalScope) => EvalValue
  }
  const parseUnary = (): ((scope: EvalScope) => EvalValue) => {
    if (atOp('-')) {
      eat()
      const inner = parseUnary()
      return (scope) => -toNumber(inner(scope))
    }
    if (atOp('+')) {
      eat()
      return parseUnary()
    }
    return parsePower()
  }
  const parsePower = (): ((scope: EvalScope) => EvalValue) => {
    const base = parsePostfix()
    if (atOp('^')) {
      eat()
      const exp = parseUnary()
      return (scope) => Math.pow(toNumber(base(scope)), toNumber(exp(scope)))
    }
    return base
  }
  const parsePostfix = (): ((scope: EvalScope) => EvalValue) => {
    let primary = parsePrimary()
    while (atOp('%')) {
      eat()
      const prev = primary
      primary = (scope) => toNumber(prev(scope)) / 100
    }
    return primary
  }
  const parsePrimary = (): ((scope: EvalScope) => EvalValue) => {
    const t = peek()
    if (!t) throw new FormulaEvalError('syntax')
    if (t.kind === 'num') {
      eat()
      const n = parseFloat(t.text)
      return () => n
    }
    if (t.kind === 'op' && t.text === '(') {
      eat()
      const inner = parseExpr()
      eat(')')
      return inner
    }
    if (t.kind === 'ident') {
      eat()
      const upper = t.text.toUpperCase()
      // function call
      if (atOp('(')) {
        eat()
        const args: Array<(scope: EvalScope) => EvalValue> = []
        if (!atOp(')')) {
          args.push(parseExpr())
          while (atOp(',')) {
            eat()
            args.push(parseExpr())
          }
        }
        eat(')')
        return (scope) => callFunction(upper, args, scope)
      }
      const dir = asDirection(t.text)
      if (dir) return (scope) => scope.direction(dir)
      if (upper === 'TRUE') return () => 1
      if (upper === 'FALSE') return () => 0
      const ref = parseCellRef(t.text)
      if (ref) {
        // range A1:B3
        if (atOp(':')) {
          eat()
          const endTok = eat()
          const endRef = parseCellRef(endTok.text)
          if (!endRef) throw new FormulaEvalError('syntax')
          return (scope) => scope.range(ref, endRef)
        }
        return (scope) => scope.cell(ref)
      }
      throw new FormulaEvalError('syntax')
    }
    throw new FormulaEvalError('syntax')
  }

  const expr = parseExpr()
  if (pos !== tokens.length) throw new FormulaEvalError('syntax')
  return expr
}

interface EvalScope {
  cell(ref: [number, number]): number
  range(from: [number, number], to: [number, number]): number[]
  direction(dir: FormulaDirection): number[]
}

function toNumber(v: EvalValue): number {
  return Array.isArray(v) ? (v[0] ?? 0) : v
}

function flatten(args: EvalValue[]): number[] {
  const out: number[] = []
  for (const a of args) {
    if (Array.isArray(a)) out.push(...a)
    else out.push(a)
  }
  return out
}

function callFunction(
  name: string,
  args: Array<(scope: EvalScope) => EvalValue>,
  scope: EvalScope,
): EvalValue {
  switch (name) {
    case 'SUM':
    case 'AVERAGE':
    case 'MIN':
    case 'MAX':
    case 'COUNT':
    case 'PRODUCT': {
      // positional operands with no numeric cells (e.g. an empty column) and
      // aggregates over no numbers produce Word's "Undefined Bookmark" result
      const nums = flatten(args.map((a) => a(scope)))
      if (nums.length === 0) throw new FormulaEvalError('empty')
      if (name === 'SUM') return nums.reduce((s, n) => s + n, 0)
      if (name === 'COUNT') return nums.length
      if (name === 'AVERAGE') return nums.reduce((s, n) => s + n, 0) / nums.length
      if (name === 'MIN') return Math.min(...nums)
      if (name === 'MAX') return Math.max(...nums)
      return nums.reduce((p, n) => p * n, 1)
    }
    case 'ROUND': {
      const [x, d] = args.map((a) => toNumber(a(scope)))
      const digits = args.length > 1 ? d : 0
      const factor = Math.pow(10, digits)
      // Word/Excel ROUND: half away from zero
      const scaled = Math.abs(x) * factor
      const rounded = Math.round(scaled + Number.EPSILON * scaled)
      return (x < 0 ? -1 : 1) * (rounded / factor)
    }
    case 'ABS':
      return Math.abs(singleArg(args, scope))
    case 'INT':
      return Math.floor(singleArg(args, scope))
    case 'SIGN': {
      const v = singleArg(args, scope)
      return v > 0 ? 1 : v < 0 ? -1 : 0
    }
    case 'MOD': {
      const [a, b] = args.map((x) => toNumber(x(scope)))
      if (b === 0) throw new FormulaEvalError('div0')
      // Excel/Word MOD: sign follows the divisor
      return a - b * Math.floor(a / b)
    }
    case 'IF': {
      if (args.length < 2 || args.length > 3) throw new FormulaEvalError('syntax')
      return toNumber(args[0](scope)) !== 0
        ? toNumber(args[1](scope))
        : args.length > 2
          ? toNumber(args[2](scope))
          : 0
    }
    case 'AND': {
      const nums = flatten(args.map((a) => a(scope)))
      return nums.every((n) => n !== 0) ? 1 : 0
    }
    case 'OR': {
      const nums = flatten(args.map((a) => a(scope)))
      return nums.some((n) => n !== 0) ? 1 : 0
    }
    case 'NOT':
      return singleArg(args, scope) === 0 ? 1 : 0
    case 'DEFINED':
      // lazy: any evaluation error inside the argument yields 0 (Word: 1/0)
      try {
        return Number.isFinite(toNumber(args[0](scope))) ? 1 : 0
      } catch {
        return 0
      }
    case 'TRUE':
      return 1
    case 'FALSE':
      return 0
    default:
      throw new FormulaEvalError('syntax')
  }
}

function singleArg(args: Array<(scope: EvalScope) => EvalValue>, scope: EvalScope): number {
  if (args.length !== 1) throw new FormulaEvalError('syntax')
  return toNumber(args[0](scope))
}

function buildScope(grid: FormulaGrid, row: number, col: number): EvalScope {
  const textAt = (r: number, c: number): string => grid.texts[r]?.[c] ?? ''
  const scope: EvalScope = {
    cell(ref) {
      const [r, c] = ref
      if (r === row && c === col) throw new FormulaEvalError('syntax') // self-reference
      // an out-of-table or empty referenced cell contributes 0, like Word
      return parseCellNumber(textAt(r, c)) ?? 0
    },
    range(from, to) {
      const out: number[] = []
      const [r1, c1] = from
      const [r2, c2] = to
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
        for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
          const n = parseCellNumber(textAt(r, c))
          if (n !== null) out.push(n)
        }
      }
      return out
    },
    direction(dir) {
      return collectDirectionOperands(grid.texts, row, col, dir, grid.isHeaderRow)
    },
  }
  return scope
}

const exprCache = new Map<string, (scope: EvalScope) => EvalValue>()

function compileExpression(expr: string): (scope: EvalScope) => EvalValue {
  let compiled = exprCache.get(expr)
  if (!compiled) {
    compiled = tokenizeAndParse(expr)
    if (exprCache.size > 256) exprCache.clear()
    exprCache.set(expr, compiled)
  }
  return compiled
}

/** general-purpose number text (no picture): trimmed float noise, no -0 */
function generalNumber(v: number): string {
  if (!Number.isFinite(v)) return '0'
  const n = Number(v.toPrecision(12))
  return Object.is(n, -0) ? '0' : String(n)
}

/**
 * Evaluate one formula field against the grid and return the display text
 * (cached result), applying the numeric picture when present. Never throws:
 * Word-parity error strings come back instead.
 */
export function evaluateFormulaInGrid(
  instr: string,
  grid: FormulaGrid,
  row: number,
  col: number,
): string {
  const parsed = parseFormulaInstruction(instr)
  if (!parsed || parsed.expr === '') return FORMULA_SYNTAX_ERROR
  try {
    const value = toNumber(compileExpression(parsed.expr)(buildScope(grid, row, col)))
    if (!Number.isFinite(value)) return FORMULA_EMPTY_ERROR
    return parsed.picture ? formatNumericPicture(value, parsed.picture) : generalNumber(value)
  } catch (e) {
    if (e instanceof FormulaEvalError) {
      return e.kind === 'div0'
        ? FORMULA_DIV_ZERO_ERROR
        : e.kind === 'empty'
          ? FORMULA_EMPTY_ERROR
          : FORMULA_SYNTAX_ERROR
    }
    throw e
  }
}

/** Word's Formula dialog prefill: the bottom of a numeric column proposes SUM(ABOVE), else SUM(LEFT) */
export function proposeTableFormula(
  texts: FormulaGridTexts,
  row: number,
  col: number,
  isHeaderRow?: HeaderRowPredicate,
): string {
  const above = collectDirectionOperands(texts, row, col, 'ABOVE', isHeaderRow)
  if (above.length > 0) return '=SUM(ABOVE)'
  const left = collectDirectionOperands(texts, row, col, 'LEFT', isHeaderRow)
  if (left.length > 0) return '=SUM(LEFT)'
  return '=SUM(ABOVE)'
}

// ---- numeric picture formatting (\# switch) ----

interface PictureSection {
  prefix: string
  suffix: string
  intZeroes: number
  decRequired: number
  decOptional: number
  grouping: boolean
  percent: boolean
}

function splitPictureSections(picture: string): string[] {
  const sections: string[] = []
  let current = ''
  let quoted = false
  for (let i = 0; i < picture.length; i++) {
    const ch = picture[i]
    if (ch === '"' && picture[i - 1] !== '\\') quoted = !quoted
    if (ch === '\\' && picture[i + 1] !== undefined) {
      current += ch + picture[i + 1]
      i++
      continue
    }
    if (ch === ';' && !quoted) {
      sections.push(current)
      current = ''
      continue
    }
    current += ch
  }
  sections.push(current)
  return sections.slice(0, 3)
}

function analyzeSection(section: string): PictureSection {
  const unescaped: Array<{ ch: string; literal: boolean }> = []
  let quoted = false
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"' && section[i - 1] !== '\\') {
      quoted = !quoted
      continue
    }
    if (ch === '\\' && section[i + 1] !== undefined) {
      unescaped.push({ ch: section[i + 1], literal: true })
      i++
      continue
    }
    unescaped.push({ ch, literal: quoted })
  }
  const firstPlaceholder = unescaped.findIndex((t) => !t.literal && '0#x?'.includes(t.ch))
  const lastPlaceholder = findLastIndex(unescaped, (t) => !t.literal && '0#x?'.includes(t.ch))
  const dotIndex = unescaped.findIndex((t) => !t.literal && t.ch === '.')
  let intZeroes = 0
  let decRequired = 0
  let decOptional = 0
  let grouping = false
  // digits before the decimal point: only '0' is required ('#' never pads);
  // after it, '0' is a required decimal and '#'/x optional ones
  unescaped.forEach((t, i) => {
    if (t.literal || firstPlaceholder === -1) return
    const afterDot = dotIndex !== -1 && i > dotIndex
    const inCore = i >= firstPlaceholder && i <= lastPlaceholder
    if (!inCore || !'0#x?'.includes(t.ch)) {
      if (!t.literal && t.ch === ',') grouping = true
      return
    }
    if (t.ch === '0') {
      if (afterDot) decRequired++
      else intZeroes++
    } else if (afterDot && (t.ch === '#' || t.ch === 'x')) {
      decOptional++
    }
  })
  return {
    prefix: unescaped
      .slice(0, firstPlaceholder === -1 ? unescaped.length : firstPlaceholder)
      .map((t) => t.ch)
      .join(''),
    // a section without any digit placeholder renders its literals only
    suffix:
      lastPlaceholder === -1
        ? ''
        : unescaped
            .slice(lastPlaceholder + 1)
            .map((t) => t.ch)
            .join(''),
    intZeroes,
    decRequired: dotIndex === -1 ? 0 : decRequired,
    decOptional: dotIndex === -1 ? 0 : decOptional,
    grouping,
    percent: unescaped.some((t) => !t.literal && t.ch === '%'),
  }
}

function findLastIndex<T>(arr: T[], pred: (item: T) => boolean): number {
  for (let i = arr.length - 1; i >= 0; i--) if (pred(arr[i])) return i
  return -1
}

function renderSection(section: string, abs: number): string {
  const p = analyzeSection(section)
  if (p.intZeroes === 0 && p.decRequired === 0 && p.decOptional === 0) {
    // literal-only section (e.g. an em-dash zero section)
    return `${p.prefix}${p.suffix}`
  }
  let v = p.percent ? abs * 100 : abs
  const decimals = p.decRequired + p.decOptional
  v = Math.round(v * Math.pow(10, decimals)) / Math.pow(10, decimals)
  const fixed = v.toFixed(decimals)
  let [int, frac = ''] = fixed.split('.')
  // optional decimals show only while significant, never below the required count
  while (frac.length > p.decRequired && frac.endsWith('0')) frac = frac.slice(0, -1)
  while (int.length < p.intZeroes) int = '0' + int
  if (p.grouping) int = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  const body = frac.length > 0 ? `${int}.${frac}` : int
  return `${p.prefix}${body}${p.suffix}`
}

/**
 * Word numeric picture switch formatting (minimal code set): 0 required digit,
 * # / x optional digits, . decimal point, , thousands grouping between digit
 * placeholders, % (value × 100), quoted or backslash-escaped literal text, and
 * up to three positive;negative;zero sections.
 */
export function formatNumericPicture(value: number, picture: string): string {
  if (!Number.isFinite(value)) return '0'
  const sections = splitPictureSections(picture)
  if (sections.length === 0) return generalNumber(value)
  let section = sections[0]
  const negative = value < 0
  if (value < 0 && sections.length > 1) section = sections[1]
  else if (value === 0 && sections.length > 2) section = sections[2]
  const body = renderSection(section, Math.abs(value))
  // a dedicated negative section spells the sign itself; the positive section
  // reused for negatives keeps the minus
  return negative && sections.length > 1 ? body : negative ? `-${body}` : body
}

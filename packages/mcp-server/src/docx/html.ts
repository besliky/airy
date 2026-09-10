// Restricted HTML <-> docx-engine block model, with no DOM dependency.
//
// The embedded agent's insert_content parses HTML with DOMParser inside the
// renderer; the headless server cannot pull that in, so this module implements
// the same restricted fragment grammar with a small hand-written parser and
// maps straight onto the engine's GeneratedBlock (paragraphs / headings /
// list items) and static w:tbl fragments (tables).
//
// Supported tags: p, h1-h6, ul, ol, li, strong/b, em/i, u, s/strike/del,
// a[href], br, blockquote, pre, code, table/tr/th/td. Everything else keeps
// its text content, mirroring the renderer's salvage rule.
import {
  TABLE_HEADER_FILL,
  decodeEntities,
  generateTableModelXml,
  type Block,
  type GeneratedBlock,
  type Run,
  type TableModel,
} from '@airy-office/docx-engine'

// ---- parsing: restricted HTML -> session content ----

/** Content produced by one insert: paragraph-ish blocks plus one table fragment. */
export type HtmlBlock = GeneratedBlock | { type: 'table'; xml: string; model: TableModel }

interface OpenEl {
  tag: string
  attrs: Record<string, string>
  parent: OpenEl | null
  children: Node[]
}

type Node = OpenEl | { text: string }

const VOID_TAGS = new Set(['br', 'hr', 'img'])

interface Token {
  kind: 'text' | 'open' | 'close'
  /** text content (kind=text), or the tag name (open/close) */
  value: string
  attrs?: Record<string, string>
  selfClosing?: boolean
}

function tokenize(html: string): Token[] {
  const tokens: Token[] = []
  let i = 0
  while (i < html.length) {
    const open = html.indexOf('<', i)
    if (open === -1) {
      pushText(tokens, html.slice(i))
      break
    }
    if (open > i) pushText(tokens, html.slice(i, open))
    // comments, doctype, CDATA: skip to the matching close
    if (html.startsWith('<!--', open)) {
      const end = html.indexOf('-->', open + 4)
      i = end === -1 ? html.length : end + 3
      continue
    }
    if (html[open + 1] === '!' || html[open + 1] === '?') {
      const end = html.indexOf('>', open + 2)
      i = end === -1 ? html.length : end + 1
      continue
    }
    const close = findTagEnd(html, open)
    if (close === -1) {
      pushText(tokens, html.slice(open))
      break
    }
    const raw = html.slice(open + 1, close)
    if (raw.startsWith('/')) {
      tokens.push({ kind: 'close', value: raw.slice(1).trim().toLowerCase() })
    } else {
      const parsed = /^([a-zA-Z][a-zA-Z0-9:-]*)([\s\S]*)$/.exec(raw)!
      const attrs = parseAttrs(parsed[2])
      const tag = parsed[1].toLowerCase()
      const selfClosing = raw.endsWith('/') || VOID_TAGS.has(tag)
      tokens.push({ kind: 'open', value: tag, attrs, selfClosing })
    }
    i = close + 1
  }
  return tokens
}

/** index of the '>' ending the tag that starts at `start`, honoring quoted attribute values */
function findTagEnd(html: string, start: number): number {
  let quote: string | null = null
  for (let i = start + 1; i < html.length; i++) {
    const ch = html[i]
    if (quote) {
      if (ch === quote) quote = null
    } else if (ch === '"' || ch === "'") {
      quote = ch
    } else if (ch === '>') {
      return i
    }
  }
  return -1
}

function parseAttrs(raw: string): Record<string, string> {
  const attrs: Record<string, string> = {}
  const re = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*(?:=\s*("([^"]*)"|'([^']*)'|[^\s"'>]+))?/g
  let m: RegExpExecArray | null
  while ((m = re.exec(raw)) !== null) {
    if (m[1]) attrs[m[1].toLowerCase()] = decodeHtmlText(m[2] ?? '')
  }
  return attrs
}

function pushText(tokens: Token[], raw: string): void {
  if (raw !== '') tokens.push({ kind: 'text', value: raw })
}

/** HTML entities: the XML five plus the two named entities models actually emit */
function decodeHtmlText(raw: string): string {
  return decodeEntities(raw.replace(/&nbsp;/gi, ' '))
}

function buildTree(tokens: Token[]): OpenEl {
  const root: OpenEl = { tag: '#root', attrs: {}, parent: null, children: [] }
  let cur = root
  for (const token of tokens) {
    if (token.kind === 'text') {
      cur.children.push({ text: decodeHtmlText(token.value) })
    } else if (token.kind === 'open') {
      // an unclosed block-level tag while a new one opens: recover by implicit
      // close (browsers auto-close <p>/<li>; we do the same for block tags)
      const el: OpenEl = { tag: token.value, attrs: token.attrs ?? {}, parent: cur, children: [] }
      cur.children.push(el)
      if (!token.selfClosing) cur = el
    } else {
      // close: walk up to the matching open tag (tolerates missing/misordered closes)
      let node: OpenEl | null = cur
      while (node && node.tag !== token.value) node = node.parent
      if (node?.parent) cur = node.parent
    }
  }
  return root
}

const BLOCK_TAGS = new Set([
  'p',
  'div',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'ul',
  'ol',
  'table',
  'blockquote',
  'pre',
  'formula',
])

interface Marks {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  href?: string
  mono?: boolean
}

const PRESET = {
  code: { font: 'Consolas', shadingFill: 'F2F2F2', borders: 'l' },
  quote: { color: '595959', indentLeft: 720, borders: 'l' },
} as const

/** inline content of an element -> Run[] (\n encodes <br>, the engine writes w:br) */
function inlineRuns(el: OpenEl, marks: Marks): Run[] {
  const runs: Run[] = []
  const push = (text: string, m: Marks) => {
    if (text === '') return
    const run: Run = { text }
    if (m.bold) run.bold = true
    if (m.italic) run.italic = true
    if (m.underline) run.underline = true
    if (m.strike) run.strike = true
    if (m.mono) run.fontAscii = PRESET.code.font
    if (m.href) run.link = { href: m.href }
    const prev = runs[runs.length - 1]
    // adjacent runs with identical formatting merge, like the editor's marks
    if (
      prev &&
      prev.bold === run.bold &&
      prev.italic === run.italic &&
      prev.underline === run.underline &&
      prev.strike === run.strike &&
      prev.fontAscii === run.fontAscii &&
      (prev.link?.href ?? null) === (run.link?.href ?? null) &&
      !prev.math &&
      !prev.image
    ) {
      prev.text += text
      return
    }
    runs.push(run)
  }
  const walk = (node: Node, m: Marks) => {
    if ('text' in node) {
      push(node.text.replace(/\s+/g, ' '), m)
      return
    }
    const tag = node.tag
    if (tag === 'br') {
      // attach the break to the current formatting level via a text \n
      push('\n', m)
      return
    }
    if (tag === 'strong' || tag === 'b') return walkChildren(node, { ...m, bold: true })
    if (tag === 'em' || tag === 'i') return walkChildren(node, { ...m, italic: true })
    if (tag === 'u') return walkChildren(node, { ...m, underline: true })
    if (tag === 's' || tag === 'strike' || tag === 'del')
      return walkChildren(node, { ...m, strike: true })
    if (tag === 'code') return walkChildren(node, { ...m, mono: true })
    if (tag === 'a') {
      const href = node.attrs['href']
      return href ? walkChildren(node, { ...m, href }) : walkChildren(node, m)
    }
    walkChildren(node, m)
  }
  const walkChildren = (node: OpenEl, m: Marks) => {
    for (const child of node.children) walk(child, m)
  }
  for (const child of el.children) walk(child, marks)
  return runs
}

function parseList(el: OpenEl, kind: 'bullet' | 'ordered', ilvl: number, out: HtmlBlock[]): void {
  for (const child of el.children) {
    if ('text' in child) continue
    if (child.tag === 'ul' || child.tag === 'ol') {
      parseList(child, child.tag === 'ol' ? 'ordered' : 'bullet', ilvl + 1, out)
      continue
    }
    if (child.tag !== 'li') continue
    // nested lists ride along as deeper items after this one
    const nested = child.children.filter(
      (c): c is OpenEl => !('text' in c) && (c.tag === 'ul' || c.tag === 'ol'),
    )
    const own: OpenEl = {
      tag: 'li',
      attrs: {},
      parent: null,
      children: child.children.filter((c) => !('text' in c) || c.text.trim() !== ''),
    }
    out.push(listItemBlock(kind, ilvl, inlineRuns(own, {})))
    for (const n of nested) parseList(n, n.tag === 'ol' ? 'ordered' : 'bullet', ilvl + 1, out)
  }
}

function listItemBlock(kind: 'bullet' | 'ordered', ilvl: number, runs: Run[]): HtmlBlock {
  // numId is a placeholder; the session rewrites it to a real numbering id
  return {
    type: 'listItem',
    list: { kind, numId: `pending-${kind}`, ilvl: Math.min(ilvl, 4) },
    runs,
  }
}

/** table element -> static w:tbl fragment (header row bold + shaded, equal columns) */
function parseTable(el: OpenEl): HtmlBlock | null {
  const rows: OpenEl[] = []
  const collectRows = (node: OpenEl) => {
    for (const child of node.children) {
      if ('text' in child) continue
      if (child.tag === 'tr') rows.push(child)
      else if (child.tag === 'thead' || child.tag === 'tbody' || child.tag === 'tfoot')
        collectRows(child)
      else if (child.tag === 'table') collectRows(child) // tolerate one nesting level
    }
  }
  collectRows(el)
  if (rows.length === 0) return null
  const grids = rows.map((tr) =>
    tr.children.filter((c): c is OpenEl => !('text' in c) && (c.tag === 'td' || c.tag === 'th')),
  )
  const cols = Math.max(...grids.map((cells) => cells.length))
  if (cols === 0) return null
  const header = grids[0].some((c) => c.tag === 'th')
  const model: TableModel = {
    rows: grids.map((cells, r) => {
      const isHeader = header && r === 0
      const row = cells.map((cell) => ({
        paras: cellParas(cell),
        ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}),
      }))
      while (row.length < cols)
        row.push({ paras: [''], ...(isHeader ? { bold: true, fill: TABLE_HEADER_FILL } : {}) })
      return row
    }),
    colWidthsPct: Array.from({ length: cols }, () => 100 / cols),
  }
  return { type: 'table', xml: generateTableModelXml(model), model }
}

/** cell paragraphs: <br> and nested block tags split paragraphs */
function cellParas(cell: OpenEl): string[] {
  const paras: string[] = []
  let current = ''
  const push = () => {
    const t = current.replace(/\s+/g, ' ').trim()
    if (t !== '') paras.push(t)
    current = ''
  }
  const walk = (node: Node) => {
    if ('text' in node) {
      current += node.text
      return
    }
    if (node.tag === 'br') {
      push()
      return
    }
    if (BLOCK_TAGS.has(node.tag)) {
      push()
      for (const child of node.children) walk(child)
      push()
      return
    }
    for (const child of node.children) walk(child)
  }
  for (const child of cell.children) walk(child)
  push()
  return paras.length > 0 ? paras : ['']
}

/**
 * Parse a restricted HTML fragment into insertable content blocks.
 * Tolerates markdown code fences and plain-text input, like the renderer's
 * parseHtmlFragment.
 */
export function parseRestrictedHtml(raw: string): HtmlBlock[] {
  let text = raw.trim()
  const fence = /```(?:html)?\s*([\s\S]*?)```/.exec(text)
  if (fence) text = fence[1].trim()
  if (!text) return []

  // plain-text response (no tags): one paragraph per blank-line-separated chunk
  if (!/<[a-z][\s\S]*>/i.test(text)) {
    return text
      .split(/\n{2,}/)
      .map((para) => decodeHtmlText(para).replace(/\s+/g, ' ').trim())
      .filter((para) => para !== '')
      .map((para) => ({ type: 'paragraph' as const, runs: [{ text: para }] }))
  }

  const root = buildTree(tokenize(text))
  const out: HtmlBlock[] = []
  for (const child of root.children) {
    if ('text' in child) {
      const t = child.text.replace(/\s+/g, ' ').trim()
      if (t) out.push({ type: 'paragraph', runs: [{ text: t }] })
      continue
    }
    const el = child
    const headingMatch = /^h([1-6])$/.exec(el.tag)
    if (headingMatch) {
      out.push({
        type: 'heading',
        level: Number(headingMatch[1]),
        runs: inlineRuns(el, {}),
      })
    } else if (el.tag === 'ul' || el.tag === 'ol') {
      parseList(el, el.tag === 'ol' ? 'ordered' : 'bullet', 0, out)
    } else if (el.tag === 'table') {
      const table = parseTable(el)
      if (table) out.push(table)
    } else if (el.tag === 'pre') {
      const code = textOf(el).replace(/^\n/, '').replace(/\s+$/, '')
      if (code) {
        const lines = code.split('\n')
        const runs: Run[] = []
        lines.forEach((line, i) => {
          if (i > 0) runs.push({ text: '\n' })
          if (line !== '') runs.push({ text: line, fontAscii: PRESET.code.font })
        })
        out.push({
          type: 'paragraph',
          format: { shadingFill: PRESET.code.shadingFill, borders: PRESET.code.borders },
          runs,
        })
      }
    } else if (el.tag === 'blockquote') {
      out.push({
        type: 'paragraph',
        format: { indentLeft: PRESET.quote.indentLeft, borders: PRESET.quote.borders },
        runs: inlineRuns(el, {}).map((run) => ({ ...run, color: PRESET.quote.color })),
      })
    } else if (el.tag === 'formula') {
      // formulas need the math pipeline; degrade to the LaTeX source as text
      const latex = textOf(el).trim()
      if (latex) out.push({ type: 'paragraph', runs: [{ text: latex }] })
    } else if (el.tag === 'p' || el.tag === 'div') {
      const runs = inlineRuns(el, {})
      if (runs.length > 0) out.push({ type: 'paragraph', runs })
    } else {
      // unknown block-ish tag: salvage the text
      const t = textOf(el).replace(/\s+/g, ' ').trim()
      if (t) out.push({ type: 'paragraph', runs: [{ text: t }] })
    }
  }
  return out
}

function textOf(el: Node): string {
  if ('text' in el) return el.text
  return el.children.map(textOf).join('')
}

// ---- serialization: session blocks -> restricted HTML (read_document) ----

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** model run -> restricted inline HTML; atomic payloads (math/images) drop to their text */
function runToHtml(run: Run): string {
  let text = escapeHtml(run.text)
  if (run.bold) text = `<strong>${text}</strong>`
  if (run.italic) text = `<em>${text}</em>`
  if (run.underline) text = `<u>${text}</u>`
  if (run.strike) text = `<s>${text}</s>`
  if (run.link?.href) text = `<a href="${escapeHtml(run.link.href)}">${text}</a>`
  return text
}

function runsToHtml(runs: Run[] | undefined): string {
  // engine run text encodes breaks as \n; restricted HTML shows them as <br>
  return (runs ?? []).map((run) => runToHtml(run).replace(/\n/g, '<br>')).join('')
}

/** engine/edited block type -> the preview list's short type name */
export function blockTypeName(block: Block | GeneratedBlock): string {
  if (block.type === 'heading') return `h${Math.min(Math.max(block.level ?? 1, 1), 9)}`
  if (block.type === 'listItem') return 'li'
  if (block.type === 'table') return 'table'
  if (block.type === 'image') return 'image'
  return 'p'
}

/** collapsed single-line preview text of a block */
export function blockPreviewText(block: Block | GeneratedBlock): string {
  if (block.type === 'paragraph' || block.type === 'heading' || block.type === 'listItem') {
    return (block.runs ?? [])
      .map((r) => r.text)
      .join('')
      .replace(/\s+/g, ' ')
      .trim()
  }
  const b = block as Block
  if (b.type === 'table') {
    const rows = b.table?.rows ?? []
    const flat = rows.map((row) => row.map((cell) => cell.paras.join(' ')).join(' | ')).join(' / ')
    return flat.replace(/\s+/g, ' ').trim()
  }
  const label = b.label ?? b.previewText
  return [label, b.previewText].filter(Boolean).join(' — ').replace(/\s+/g, ' ').trim()
}

/** one block as restricted HTML (list grouping is handled by blocksToHtml) */
export function blockToHtml(block: Block | GeneratedBlock): string {
  if (block.type === 'heading') {
    const level = Math.min(Math.max(block.level ?? 1, 1), 6)
    return `<h${level}>${runsToHtml(block.runs)}</h${level}>`
  }
  if (block.type === 'listItem') {
    return `<li>${runsToHtml(block.runs)}</li>` // wrapped by blocksToHtml
  }
  if (block.type === 'table') {
    const rows = (block as Block).table?.rows ?? []
    const header = rows.length > 1 && rows[0].every((c) => c.bold || c.fill !== undefined)
    const trs = rows.map((row, r) => {
      const tag = header && r === 0 ? 'th' : 'td'
      return `<tr>${row.map((cell) => `<${tag}>${escapeHtml(cell.paras.join('\n'))}</${tag}>`).join('')}</tr>`
    })
    return `<table>${trs.join('')}</table>`
  }
  if (block.type === 'image') {
    const b = block as Block
    const preview = b.previewText ?? ''
    return `<p>[Protected image${preview ? `: ${escapeHtml(preview)}` : ''}, kept as is]</p>`
  }
  if (block.type === 'passthrough') {
    const b = block as Block
    const label = b.label ?? 'content'
    const preview = (b.previewText ?? '').replace(/\s+/g, ' ').trim()
    return `<p>[Protected content: ${escapeHtml(String(label))}${preview ? ` — visible text: "${escapeHtml(preview)}"` : ''}, kept as is]</p>`
  }
  return `<p>${runsToHtml(block.runs)}</p>`
}

/** serialize a block range to restricted HTML, grouping consecutive list items */
export function blocksToHtml(blocks: Array<Block | GeneratedBlock>): string {
  const parts: string[] = []
  let list: { kind: string; items: string[] } | null = null
  const flush = () => {
    if (!list) return
    parts.push(
      `<${list.kind === 'ordered' ? 'ol' : 'ul'}>${list.items.join('')}</${list.kind === 'ordered' ? 'ol' : 'ul'}>`,
    )
    list = null
  }
  for (const block of blocks) {
    if (block.type === 'listItem') {
      const kind = block.list?.kind ?? 'bullet'
      if (!list || list.kind !== kind) {
        flush()
        list = { kind, items: [] }
      }
      list.items.push(blockToHtml(block))
    } else {
      flush()
      parts.push(blockToHtml(block))
    }
  }
  flush()
  return parts.join('\n')
}

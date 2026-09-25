import {
  decodeEntities,
  type Block,
  type Run,
  type TableCell,
  type TableModel,
  type TableParagraph,
} from '@airy-office/docx-engine'

/**
 * Recovery for malformed documents whose table rows sit loose at body level
 * (a <w:tbl> wrapper lost to a broken producer): the parse degrades every
 * <w:tr> into a passthrough block, which renders as an opaque chip — the
 * whole table is invisible and uneditable with zero affordance (BUG-1707).
 *
 * Consecutive loose-row blocks are re-synthesized into a native table block
 * through a minimal OOXML subset (text runs, w:fldSimple table formulas,
 * gridSpan/vMerge/tcW basics); anything the subset does not model degrades to
 * plain text. The merged block is editor-native, so it renders and edits like
 * any table, and a save regenerates a single valid <w:tbl> (repairing the
 * malformed body). Valid w:tbl tables never take this path.
 */

interface Element {
  name: string
  /** full element source, open tag through close tag */
  xml: string
  /** attributes of the open tag, decoded */
  attrs: Record<string, string>
  /** inner source between open and close tags */
  inner: string
}

// recovery-grade tag scanner: the unquoted-attr branch excludes '/', so a
// self-closing tag's slash is always captured by the (\/?) group and depth
// tracking sees <w:tcW …/> as the empty element it is
const TAG_RE = /<(\/?)([\w.-]+:[\w.-]+)((?:"[^"]*"|'[^']*'|[^"'>/])*)(\/?)>/g

/** attribute map of an open-tag source */
function parseAttrs(source: string): Record<string, string> {
  const attrs: Record<string, string> = {}
  for (const m of source.matchAll(/([\w.-]+:[\w.-]+)\s*=\s*"([^"]*)"/g)) {
    attrs[m[1]] = decodeEntities(m[2] ?? '')
  }
  return attrs
}

/** top-level child elements of an element-inner source (recovery-grade scan) */
function childElements(inner: string): Element[] {
  const out: Element[] = []
  TAG_RE.lastIndex = 0
  let depth = 0
  let start = -1
  let name = ''
  let attrs: Record<string, string> = {}
  for (const m of inner.matchAll(TAG_RE)) {
    const [, closing, tag, attrSource, selfClose] = m
    if (closing) {
      depth--
      if (depth === 0 && start >= 0) {
        const xml = inner.slice(start, (m.index ?? 0) + m[0].length)
        out.push({
          name,
          xml,
          attrs,
          inner: xml.slice(xml.indexOf('>') + 1, xml.lastIndexOf('</')),
        })
        start = -1
      }
    } else if (depth === 0) {
      start = m.index ?? 0
      name = tag
      attrs = parseAttrs(attrSource ?? '')
      if (selfClose === '/') {
        const xml = inner.slice(start, (m.index ?? 0) + m[0].length)
        out.push({ name, xml, attrs, inner: '' })
        start = -1
      } else {
        depth++
      }
    } else if (selfClose !== '/') {
      depth++
    }
  }
  return out
}

function textOf(element: Element): string {
  return decodeEntities(
    [...element.xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
      .map((m) => m[1] ?? '')
      .join(''),
  )
}

/** truthy w:val ("1"/"true"/"on") or a bare element; "0"/"false" is off */
function valOn(attrs: Record<string, string>): boolean {
  const v = attrs['w:val']
  return v === undefined || !(v === '0' || v === 'false' || v === 'off')
}

function runsOf(inner: string): Run[] {
  const runs: Run[] = []
  for (const child of childElements(inner)) {
    if (child.name === 'w:r') {
      let bold = false
      let color: string | undefined
      for (const part of childElements(child.inner)) {
        if (part.name === 'w:rPr') {
          for (const rpr of childElements(part.inner)) {
            if (rpr.name === 'w:b' && valOn(rpr.attrs)) bold = true
            if (rpr.name === 'w:color' && rpr.attrs['w:val'] && rpr.attrs['w:val'] !== 'auto')
              color = rpr.attrs['w:val']
          }
        }
      }
      for (const part of childElements(child.inner)) {
        if (part.name === 'w:t') {
          const text = decodeEntities(part.inner)
          if (text) runs.push({ text, ...(bold && { bold }), ...(color && { color }) })
        } else if (part.name === 'w:br') {
          runs.push({ text: '\n' })
        } else if (part.name === 'w:tab') {
          runs.push({ text: '\t' })
        }
      }
    } else if (child.name === 'w:fldSimple') {
      // table formula field: keep the instruction on the run so the editor's
      // tableFormula mark and the save-side fldSimple regeneration apply
      const instr = (child.attrs['w:instr'] ?? '').trim()
      const cached = textOf(child) || ' '
      if (instr.startsWith('=')) {
        runs.push({ text: cached, formulaField: instr })
      } else {
        runs.push({ text: cached })
      }
    } else if (child.name === 'w:hyperlink') {
      runs.push(...runsOf(child.inner))
    }
  }
  return runs
}

function paragraphsOf(inner: string): TableParagraph[] {
  const paras: TableParagraph[] = []
  for (const child of childElements(inner)) {
    if (child.name === 'w:p') paras.push({ runs: runsOf(child.inner) })
  }
  return paras
}

function parseLooseCell(tc: Element): { cell: TableCell; widthTwips?: number } {
  let colSpan = 1
  let vMerge: TableCell['vMerge']
  let widthTwips: number | undefined
  for (const child of childElements(tc.inner)) {
    if (child.name !== 'w:tcPr') continue
    for (const pr of childElements(child.inner)) {
      if (pr.name === 'w:tcW' && (!pr.attrs['w:type'] || pr.attrs['w:type'] === 'dxa')) {
        const w = Number(pr.attrs['w:w'])
        if (w > 0) widthTwips = w
      } else if (pr.name === 'w:gridSpan') {
        colSpan = Math.max(1, Number(pr.attrs['w:val']) || 1)
      } else if (pr.name === 'w:vMerge') {
        vMerge = pr.attrs['w:val'] === 'restart' ? 'restart' : 'continue'
      }
    }
  }
  const paras = paragraphsOf(tc.inner)
  // nested tables degrade to their text: the recovery subset keeps content
  // visible, full nested fidelity stays the valid-table parse's job
  for (const child of childElements(tc.inner)) {
    if (child.name === 'w:tbl') {
      const text = textOf(child)
      if (text) paras.push({ runs: [{ text }] })
    }
  }
  return {
    cell: {
      paras: paras.map((p) => p.runs.map((r) => r.text).join('')),
      ...(paras.length > 0 ? { richParas: paras } : { richParas: [{ runs: [] }] }),
      ...(colSpan > 1 && { colSpan }),
      ...(vMerge && { vMerge }),
    },
    widthTwips,
  }
}

interface LooseRow {
  cells: TableCell[]
  /** per-cell dxa w:tcW (span cells carry none), aligned with cells */
  widths: Array<number | undefined>
  repeatHeader: boolean
}

function parseLooseRow(tr: Element): LooseRow {
  let repeatHeader = false
  const cells: TableCell[] = []
  const widths: Array<number | undefined> = []
  for (const child of childElements(tr.inner)) {
    if (child.name === 'w:trPr') {
      repeatHeader = childElements(child.inner).some((pr) => pr.name === 'w:tblHeader')
    } else if (child.name === 'w:tc') {
      const { cell, widthTwips } = parseLooseCell(child)
      cells.push(cell)
      widths.push(widthTwips)
    }
  }
  return { cells, widths, repeatHeader }
}

/** table model of loose row XML fragments; null when no row yields cells */
export function looseRowsTableModel(rowXmls: string[]): TableModel | null {
  const rows: LooseRow[] = []
  for (const xml of rowXmls) {
    const openEnd = xml.indexOf('>')
    const closeStart = xml.lastIndexOf('</w:tr>')
    if (openEnd < 0 || closeStart < 0) continue
    const row = parseLooseRow({
      name: 'w:tr',
      xml,
      attrs: {},
      inner: xml.slice(openEnd + 1, closeStart),
    })
    if (row.cells.length === 0) continue
    rows.push(row)
  }
  if (rows.length === 0) return null
  const cells = rows.map((row) => row.cells)
  // column widths from the widest dxa tcW per grid slot; only fully known
  // grids declare widths (partial data would skew the ratio worse than autofit)
  const widths: number[] = []
  for (const row of rows) {
    let col = 0
    for (let i = 0; i < row.cells.length; i++) {
      const span = row.cells[i].colSpan ?? 1
      const w = row.widths[i]
      if (w !== undefined && span === 1) widths[col] = Math.max(widths[col] ?? 0, w)
      col += span
    }
  }
  const colCount = Math.max(
    ...cells.map((row) => row.reduce((sum, c) => sum + (c.colSpan ?? 1), 0)),
  )
  // explicit index scan: widths is sparse and Array.every skips holes
  let complete = true
  for (let c = 0; c < colCount; c++) {
    if (!((widths[c] ?? 0) > 0)) complete = false
  }
  const repeatHeaderRows = rows.map((row) => (row.repeatHeader ? true : null))
  return {
    rows: cells,
    autoLayout: true,
    ...(complete && { colWidthsTwips: widths.slice(0, colCount) }),
    ...(repeatHeaderRows.some((r) => r === true) && { repeatHeaderRows }),
  }
}

/** a passthrough block that is one loose body-level table row */
export function isLooseTableRowBlock(block: Block): boolean {
  return (
    block.type === 'passthrough' &&
    block.label === 'w:tr' &&
    typeof block.originalXml === 'string' &&
    block.originalXml.includes('<w:tc')
  )
}

/**
 * Collapse each maximal run of loose-row passthrough blocks into one native
 * table block (its docxIndex anchors to the first row, so section budgets
 * still apply; the row anchors themselves are not consumed on save — the
 * merged table regenerates as a fresh valid <w:tbl>).
 */
export function mergeLooseTableRowBlocks(blocks: Block[]): Block[] {
  if (!blocks.some((b) => isLooseTableRowBlock(b))) return blocks
  const out: Block[] = []
  let run: Block[] = []
  const flush = () => {
    if (run.length === 0) return
    const model = looseRowsTableModel(run.map((b) => b.originalXml ?? ''))
    const anchor = run[0]
    if (model) {
      out.push({
        id: `${anchor.id}-loose-tbl`,
        type: 'table',
        docxIndex: anchor.docxIndex,
        originalXml: null,
        label: `Table ${model.rows.length}×${model.rows[0].length}`,
        previewText: model.rows
          .flat()
          .map((c) => c.paras.join(' '))
          .join(' ')
          .slice(0, 120),
        table: model,
      })
    } else {
      out.push(...run)
    }
    run = []
  }
  for (const block of blocks) {
    if (isLooseTableRowBlock(block)) {
      run.push(block)
    } else {
      flush()
      out.push(block)
    }
  }
  flush()
  return out
}

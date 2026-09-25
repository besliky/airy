/**
 * Chart embedded workbook (ppt/embeddings/*.xlsx) — the "Edit Data" sheet.
 *
 * PowerPoint stores a chart's data in an embedded xlsx referenced from the
 * chart part's own rels (relationship type …/package) and pointed at by
 * <c:externalData> in chartN.xml. App-inserted charts previously shipped
 * without one ("Edit Data" unavailable, see chart-insert.ts); foreign charts
 * carry one but it went stale on every edit. Both now stay in sync: on insert
 * a workbook is built from the chart data, on edit the Sheet1 data rectangle
 * is rewritten (same layout as PowerPoint: A1 empty, B1.. series names,
 * A2.. categories, B2.. values) so the numbers PowerPoint itself shows when
 * the user clicks "Edit Data" always match the rendered chart.
 *
 * The container work is synchronous on purpose: chart edits run inside the
 * guided-op executor (apply() is sync) and snapshot-based undo copies the
 * archive entry map, so the workbook entry must be replaced in the same
 * turn as the chart part. JSZip only inflates asynchronously, so a minimal
 * central-directory zip reader/writer over node:zlib lives here instead
 * (chart embeddings are tiny single-digit-part containers; zip64 is
 * explicitly bailed out of).
 */
import { inflateRawSync, deflateRawSync } from 'node:zlib'
import { parseChartXml, type ChartModel } from './chart'
import type { PackageArchive, Relationship } from './zip'
import { relsPathFor, resolveTarget } from './zip'
import { escapeXmlText } from './xml-utils'
import type { Slide } from './types'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

/** spreadsheetml package relationship: chart part rels → embedded workbook */
const PACKAGE_REL_SUFFIX = '/package'
/** [Content_Types].xml Default for the .xlsx extension */
const XLSX_DEFAULT =
  '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'

/** The chart data as laid out in the embedded Sheet1 (header row + data rows). */
export interface ChartWorkbookTable {
  categories: string[]
  series: Array<{ name: string; values: Array<number | null> }>
}

// ── minimal zip container (read + write, stored/deflate only) ────────────

const EOCD_SIG = 0x06054b50
const CEN_SIG = 0x02014b50
const LOC_SIG = 0x04034b50
/** entry count/count fields holding this value mean zip64 — not supported here */
const ZIP64_MARKER = 0xffff
const ZIP64_SIZE_MARKER = 0xffffffff

/** CRC-32 (IEEE 802.3), zip's required uncompressed-data checksum */
const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crc32(data: Uint8Array): number {
  let c = 0xffffffff
  for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/**
 * Read every entry of a zip container (central-directory driven, so data
 * descriptors and per-entry extra fields are handled by construction).
 * Returns null for anything this minimal reader refuses: zip64 markers,
 * compression methods other than stored/deflate, truncated directories.
 */
export function readZipContainer(bytes: Uint8Array): Map<string, Uint8Array> | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // locate the End Of Central Directory record: scan back over the max legal
  // comment (64 KiB) plus the fixed record size
  const minEocd = Math.max(0, bytes.length - 22 - 0xffff)
  let eocd = -1
  for (let i = bytes.length - 22; i >= minEocd; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) {
      eocd = i
      break
    }
  }
  if (eocd < 0) return null
  const count = view.getUint16(eocd + 10, true)
  let cdOffset = view.getUint32(eocd + 16, true)
  if (count === ZIP64_MARKER || cdOffset === ZIP64_SIZE_MARKER) return null
  if (cdOffset + 4 > bytes.length) return null

  const out = new Map<string, Uint8Array>()
  for (let n = 0; n < count; n++) {
    if (cdOffset + 46 > bytes.length || view.getUint32(cdOffset, true) !== CEN_SIG) return null
    const method = view.getUint16(cdOffset + 10, true)
    const csize = view.getUint32(cdOffset + 20, true)
    const nameLen = view.getUint16(cdOffset + 28, true)
    const extraLen = view.getUint16(cdOffset + 30, true)
    const commentLen = view.getUint16(cdOffset + 32, true)
    const localOffset = view.getUint32(cdOffset + 42, true)
    if (csize === ZIP64_SIZE_MARKER) return null
    const name = Buffer.from(bytes.subarray(cdOffset + 46, cdOffset + 46 + nameLen)).toString(
      'utf8',
    )
    cdOffset += 46 + nameLen + extraLen + commentLen
    // local header: fixed 30 bytes + its own name/extra lengths (they may
    // differ from the central copies)
    if (localOffset + 30 > bytes.length || view.getUint32(localOffset, true) !== LOC_SIG)
      return null
    const lNameLen = view.getUint16(localOffset + 26, true)
    const lExtraLen = view.getUint16(localOffset + 28, true)
    const dataStart = localOffset + 30 + lNameLen + lExtraLen
    if (dataStart + csize > bytes.length) return null
    const raw = bytes.subarray(dataStart, dataStart + csize)
    let data: Uint8Array
    if (method === 0) {
      data = raw
    } else if (method === 8) {
      try {
        data = inflateRawSync(raw)
      } catch {
        return null
      }
    } else {
      return null
    }
    out.set(name, data)
  }
  return out
}

/** one container entry for writeZipContainer */
export interface ZipPart {
  name: string
  data: Uint8Array
}

/** Build a deflate-compressed zip container (local headers + central directory + EOCD). */
export function writeZipContainer(parts: readonly ZipPart[]): Uint8Array {
  const locals: Buffer[] = []
  const centrals: Buffer[] = []
  let offset = 0
  for (const part of parts) {
    const name = Buffer.from(part.name, 'utf8')
    const compressed = deflateRawSync(part.data, { level: 6 })
    const crc = crc32(part.data)
    const local = Buffer.alloc(30 + name.length)
    local.writeUInt32LE(LOC_SIG, 0)
    local.writeUInt16LE(20, 4) // version needed
    local.writeUInt16LE(0, 6) // flags
    local.writeUInt16LE(8, 8) // method: deflate
    local.writeUInt16LE(0, 10) // mod time
    local.writeUInt16LE(0x21, 12) // mod date: 1980-01-01
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(compressed.length, 18)
    local.writeUInt32LE(part.data.length, 22)
    local.writeUInt16LE(name.length, 26)
    local.writeUInt16LE(0, 28) // extra len
    name.copy(local, 30)
    locals.push(local, Buffer.from(compressed))

    const central = Buffer.alloc(46 + name.length)
    central.writeUInt32LE(CEN_SIG, 0)
    central.writeUInt16LE(20, 4) // version made by
    central.writeUInt16LE(20, 6) // version needed
    central.writeUInt16LE(0, 8) // flags
    central.writeUInt16LE(8, 10) // method
    central.writeUInt16LE(0, 12) // mod time
    central.writeUInt16LE(0x21, 14) // mod date
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(compressed.length, 20)
    central.writeUInt32LE(part.data.length, 24)
    central.writeUInt16LE(name.length, 28)
    central.writeUInt16LE(0, 30) // extra len
    central.writeUInt16LE(0, 32) // comment len
    central.writeUInt16LE(0, 34) // disk start
    central.writeUInt16LE(0, 36) // internal attrs
    central.writeUInt32LE(0, 38) // external attrs
    central.writeUInt32LE(offset, 42) // local header offset
    name.copy(central, 46)
    centrals.push(central)
    offset += local.length + compressed.length
  }
  const cdSize = centrals.reduce((n, b) => n + b.length, 0)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(EOCD_SIG, 0)
  eocd.writeUInt16LE(0, 4)
  eocd.writeUInt16LE(0, 6)
  eocd.writeUInt16LE(parts.length, 8)
  eocd.writeUInt16LE(parts.length, 10)
  eocd.writeUInt32LE(cdSize, 12)
  eocd.writeUInt32LE(offset, 16) // cd offset
  eocd.writeUInt16LE(0, 20) // comment len
  return Buffer.concat([...locals, ...centrals, eocd])
}

// ── Sheet1 XML (build + parse + patch) ───────────────────────────────────

/** Excel column letters: A, B, … Z, AA … */
const xlsxColLetter = (i: number): string => {
  let s = ''
  let n = i + 1
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** column letters (A, B, … AA) → 1-based column index */
const xlsxColIndex = (letters: string): number => {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

/** numeric category text → number; null for labels (scatter x values etc.) */
function catNum(v: string): number | null {
  const t = v.trim()
  if (t === '') return null
  const n = Number(t)
  return Number.isFinite(n) ? n : null
}

/** one worksheet cell as raw XML plus its 1-based column index (0 = unknown) */
interface SheetCell {
  col: number
  xml: string
}

function parseSheetCells(cellsXml: string): SheetCell[] {
  const out: SheetCell[] = []
  const re = /<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(cellsXml)) !== null) {
    const letters = /\br="([A-Z]+)\d+"/.exec(m[0])?.[1]
    out.push({ col: letters ? xlsxColIndex(letters) : 0, xml: m[0] })
  }
  return out
}

/**
 * Build a minimal but valid xlsx container holding one Sheet1 with the chart
 * data (header row + data rows). Layout matches PowerPoint's own embeddings:
 *   A1        | B1 (ser 0 name) | C1 (ser 1 name) …
 *   A2 (cat0) | B2 (val 0,0)    | C2 (val 1,0)   …
 */
export function buildChartWorkbookBytes(table: ChartWorkbookTable): Uint8Array {
  const rows = table.categories.length
  const serCount = table.series.length

  // shared strings: all text cells collected in order
  const strings: string[] = []
  const si = (s: string): number => {
    const idx = strings.indexOf(s)
    if (idx !== -1) return idx
    strings.push(s)
    return strings.length - 1
  }

  const headerCells: string[] = []
  headerCells.push(`<c r="A1" t="s"><v>${si('')}</v></c>`)
  for (let j = 0; j < serCount; j++) {
    headerCells.push(
      `<c r="${xlsxColLetter(j + 1)}1" t="s"><v>${si(table.series[j]!.name)}</v></c>`,
    )
  }
  const dataRows: string[] = []
  for (let i = 0; i < rows; i++) {
    const rowNum = i + 2
    const cells: string[] = []
    // numeric categories (scatter x values, numeric labels) become number cells
    const xNum = catNum(table.categories[i]!)
    cells.push(
      xNum !== null
        ? `<c r="A${rowNum}"><v>${xNum}</v></c>`
        : `<c r="A${rowNum}" t="s"><v>${si(table.categories[i]!)}</v></c>`,
    )
    for (let j = 0; j < serCount; j++) {
      const val = table.series[j]!.values[i]
      if (val !== null && val !== undefined) {
        cells.push(`<c r="${xlsxColLetter(j + 1)}${rowNum}"><v>${val}</v></c>`)
      }
    }
    dataRows.push(`<row r="${rowNum}">${cells.join('')}</row>`)
  }

  const sheetXml =
    XML_DECL +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="A1:${xlsxColLetter(serCount)}${rows + 1}"/>` +
    '<sheetData>' +
    `<row r="1">${headerCells.join('')}</row>` +
    dataRows.join('') +
    '</sheetData></worksheet>'

  const sharedStringsXml =
    XML_DECL +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
    strings.map((s) => `<si><t>${escapeXmlText(s)}</t></si>`).join('') +
    '</sst>'

  const workbookXml =
    XML_DECL +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'

  const workbookRels =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '</Relationships>'

  const topRels =
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>'

  const contentTypes =
    XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '</Types>'

  return writeZipContainer([
    { name: '[Content_Types].xml', data: Buffer.from(contentTypes, 'utf8') },
    { name: '_rels/.rels', data: Buffer.from(topRels, 'utf8') },
    { name: 'xl/workbook.xml', data: Buffer.from(workbookXml, 'utf8') },
    { name: 'xl/_rels/workbook.xml.rels', data: Buffer.from(workbookRels, 'utf8') },
    { name: 'xl/worksheets/sheet1.xml', data: Buffer.from(sheetXml, 'utf8') },
    { name: 'xl/sharedStrings.xml', data: Buffer.from(sharedStringsXml, 'utf8') },
  ])
}

/**
 * Read the chart data rectangle out of an embedded xlsx: series names from the
 * header row (B1..), categories from column A (A2..), values from the body.
 * Returns null when the container or its Sheet1 is unreadable.
 */
export function readChartWorkbookBytes(bytes: Uint8Array): ChartWorkbookTable | null {
  const parts = readZipContainer(bytes)
  if (!parts) return null
  const sheet = parts.get('xl/worksheets/sheet1.xml')
  if (!sheet) return null
  const sheetXml = Buffer.from(sheet).toString('utf8')
  if (!/<sheetData\/>|<sheetData[\s>]/.test(sheetXml)) return null

  // shared strings (t="s" cells index into the sst)
  const shared: string[] = []
  const sst = parts.get('xl/sharedStrings.xml')
  if (sst) {
    const sstXml = Buffer.from(sst).toString('utf8')
    for (const m of sstXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
      shared.push(
        [...m[1]!.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]!).join('') ?? '',
      )
    }
  }

  // decode a cell's text/value: t="s" shared string, t="inlineStr" inline text,
  // t="str" formula string result, default numeric
  const cellValue = (cellXml: string): string | number | null => {
    const t = /\bt="([^"]+)"/.exec(cellXml)?.[1] ?? ''
    if (t === 'inlineStr') {
      const text = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(cellXml)?.[1]
      return text === undefined ? null : text
    }
    const v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cellXml)?.[1]
    if (v === undefined) return null
    if (t === 's') {
      const idx = parseInt(v, 10)
      return Number.isInteger(idx) ? (shared[idx] ?? '') : ''
    }
    const n = Number(v)
    return Number.isFinite(n) ? n : v
  }

  // grid[row][col] of decoded cell values
  const grid = new Map<number, Map<number, string | number | null>>()
  let maxRow = 0
  let maxCol = 0
  for (const rm of sheetXml.matchAll(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g)) {
    const rowXml = rm[0]
    const rNum = Number(/\br="(\d+)"/.exec(rowXml)?.[1] ?? 0)
    if (!rNum) continue
    const row = new Map<number, string | number | null>()
    for (const cell of parseSheetCells(
      rowXml.slice(rowXml.indexOf('>') + 1, rowXml.length - '</row>'.length),
    )) {
      if (cell.col < 1) continue // cell without an r= reference: not addressable
      row.set(cell.col, cellValue(cell.xml))
    }
    grid.set(rNum, row)
    maxRow = Math.max(maxRow, rNum)
    maxCol = Math.max(maxCol, ...row.keys())
  }

  // series columns: scan rightward from B; a column joins when its row-1
  // header CELL exists (empty names are legal series names) or it carries
  // values. A gap of ≥2 empty columns ends the scan — helper/notes columns
  // further right are workbook content, not chart series, and must never be
  // adopted into the chart on an Edit Data round-trip.
  const header = grid.get(1)
  const colHasValue = (col: number): boolean => {
    for (let r = 2; r <= maxRow; r++) {
      const v = grid.get(r)?.get(col)
      if (v != null && String(v).trim() !== '') return true
    }
    return false
  }
  const seriesCols: number[] = []
  let misses = 0
  for (let col = 2; col <= maxCol; col++) {
    if ((header?.has(col) ?? false) || colHasValue(col)) {
      seriesCols.push(col)
      misses = 0
    } else if (++misses >= 2) {
      break
    }
  }
  const series = seriesCols.map((col) => ({
    name: String(header?.get(col) ?? ''),
    values: [] as Array<number | null>,
  }))

  // data rows: column A is the category column, the series columns carry values
  const categories: string[] = []
  for (let r = 2; r <= maxRow; r++) {
    const row = grid.get(r)
    const catRaw = row?.get(1)
    const vals = seriesCols.map((col) => {
      const raw = row?.get(col)
      if (raw == null) return null
      const n = typeof raw === 'number' ? raw : Number(raw)
      return Number.isFinite(n) && String(raw).trim() !== '' ? n : null
    })
    // a row is part of the rectangle when it has a category or any value
    const hasCat = catRaw != null && String(catRaw).trim() !== ''
    const hasVal = vals.some((v) => v !== null)
    if (!hasCat && !hasVal) continue
    categories.push(catRaw == null ? '' : String(catRaw))
    vals.forEach((v, i) => series[i]!.values.push(v))
  }

  return { categories, series }
}

/** patch the Sheet1 data rectangle inside an embedded xlsx: rewrite only the
 * cells of A1:last (header + category column + series columns) and leave every
 * other cell untouched, so Excel-authored helpers, notes and formulas outside
 * the chart range survive. Inside the rectangle a cell that carried a <f>
 * formula is left byte-identical: its cached <v> is the workbook's own recalc
 * result, and rewriting it would be rolled back by the first Edit-Data recalc.
 * Returns the updated container bytes, or null on failure. */
export function patchChartWorkbookBytes(
  bytes: Uint8Array,
  table: ChartWorkbookTable,
): Uint8Array | null {
  const parts = readZipContainer(bytes)
  const sheetPart = parts?.get('xl/worksheets/sheet1.xml')
  if (!parts || !sheetPart) return null
  const sheetXml = Buffer.from(sheetPart).toString('utf8')
  if (!/<sheetData\/>|<sheetData[\s>]/.test(sheetXml)) return null

  // inline strings keep sharedStrings.xml (and every other part) valid
  // without touching it
  const esc = escapeXmlText
  const inlineStrCell = (ref: string, text: string, style = '') =>
    `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`
  const numCell = (ref: string, val: number, style = '') =>
    `<c r="${ref}"${style}><v>${val}</v></c>`

  // existing rows by number, verbatim xml + parsed cells
  const existingRows = new Map<number, { xml: string; open: string; cells: SheetCell[] }>()
  const rowRe = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g
  let maxExistingRow = 0
  let maxExistingCol = 0
  let rm: RegExpExecArray | null
  while ((rm = rowRe.exec(sheetXml)) !== null) {
    const rowXml = rm[0]
    const rNum = Number(/\br="(\d+)"/.exec(rowXml)?.[1] ?? 0)
    if (!rNum) continue
    // every row counts, cell-less ones too: a trailing <row> that only
    // carries ht/hidden/style attrs must survive the rewrite
    maxExistingRow = Math.max(maxExistingRow, rNum)
    const selfClosing = /\/>$/.test(rowXml)
    const open = selfClosing ? rowXml.slice(0, -2) + '>' : rowXml.slice(0, rowXml.indexOf('>') + 1)
    const cells = parseSheetCells(
      selfClosing ? '' : rowXml.slice(open.length, rowXml.length - '</row>'.length),
    )
    for (const c of cells) maxExistingCol = Math.max(maxExistingCol, c.col)
    existingRows.set(rNum, { xml: rowXml, open, cells })
  }

  const lastCol = table.series.length + 1 // A (categories) + one column per series
  const lastDataRow = table.categories.length + 1

  /** new rectangle cell for one position, merging with the existing cell */
  const mergedCell = (
    col: number,
    row: number,
    value: string | number | null,
  ): { col: number; xml: string } | null => {
    const existing = existingRows.get(row)?.cells.find((c) => c.col === col)
    if (value === null) return null // point removed: the old cell goes too
    if (existing && /<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>)/.test(existing.xml)) {
      // an authored formula owns this cell — its cached <v> is the workbook's
      // own recalc result, not the chart's display value. The cell stays
      // byte-identical; the chart part's numCache carries the edited value.
      return { col, xml: existing.xml }
    }
    const style =
      existing && /\bs="\d+"/.test(existing.xml) ? ` ${/\bs="\d+"/.exec(existing.xml)![0]}` : ''
    const ref = `${xlsxColLetter(col - 1)}${row}` // col is 1-based, letters 0-based
    return {
      col,
      xml:
        typeof value === 'number' ? numCell(ref, value, style) : inlineStrCell(ref, value, style),
    }
  }

  const rowsXml: string[] = []
  for (let r = 1; r <= Math.max(lastDataRow, maxExistingRow); r++) {
    const existing = existingRows.get(r)
    if (r > lastDataRow) {
      if (existing) rowsXml.push(existing.xml) // outside the rectangle: byte-identical
      continue
    }
    // rectangle row: header first, then category + series values
    const values: Array<string | number | null> =
      r === 1
        ? ['', ...table.series.map((s) => s.name)]
        : [
            catNum(table.categories[r - 2]!) ?? table.categories[r - 2]!,
            ...table.series.map((s) => s.values[r - 2] ?? null),
          ]
    const cells: Array<{ col: number; xml: string }> = []
    for (let col = 1; col <= lastCol; col++) {
      const cell = mergedCell(col, r, values[col - 1] ?? null)
      if (cell) cells.push(cell)
    }
    // cells right of the data rectangle survive (helper columns, notes)
    for (const cell of existing?.cells ?? []) {
      if (cell.col > lastCol) cells.push(cell)
    }
    cells.sort((a, b) => a.col - b.col)
    if (cells.length === 0 && !existing) continue
    rowsXml.push(
      `${existing ? existing.open : `<row r="${r}">`}${cells.map((c) => c.xml).join('')}</row>`,
    )
  }
  const newSheetData = `<sheetData>${rowsXml.join('')}</sheetData>`

  let updatedSheet = sheetXml.replace(
    /<sheetData\/>|<sheetData[^>]*>[\s\S]*?<\/sheetData>/,
    newSheetData,
  )
  // dimension covers the union of the old extent and the new data rectangle;
  // all column indexes here are 1-based
  const dim = /<dimension ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"\s*\/>/.exec(sheetXml)
  const dimCol = dim ? xlsxColIndex(dim[3] ?? dim[1]) : 1
  const dimRow = dim ? Number(dim[4] ?? dim[2]) : 1
  const lastRefCol = Math.max(lastCol, dimCol, maxExistingCol)
  const lastRef = `${xlsxColLetter(lastRefCol - 1)}${Math.max(lastDataRow, dimRow, maxExistingRow)}`
  updatedSheet = updatedSheet.replace(/<dimension[^>]*\/>/, `<dimension ref="A1:${lastRef}"/>`)

  const out: ZipPart[] = []
  for (const [name, data] of parts) {
    out.push({
      name,
      data: name === 'xl/worksheets/sheet1.xml' ? Buffer.from(updatedSheet, 'utf8') : data,
    })
  }
  return writeZipContainer(out)
}

// ── package wiring (chart part rels / externalData / Content_Types) ──────

/**
 * Find the chart part a chart element references (slide rels → r:id →
 * ppt/charts/chartN.xml). Returns null for non-chart elements or dangling refs.
 */
export function findChartPartPath(
  archive: PackageArchive,
  slide: Slide,
  elementId: string,
): string | null {
  const el = slide.elements.find((e) => e.id === elementId)
  if (!el || el.type !== 'chart') return null
  const xml = (el as { anchor?: { originalXml?: string } }).anchor?.originalXml
  const rId = xml ? /r:id="([^"]+)"/.exec(xml)?.[1] : undefined
  if (!rId) return null
  for (const rel of archive.readRels(slide.path).values()) {
    if (rel.id === rId && rel.type.endsWith('/chart')) return resolveTarget(slide.path, rel.target)
  }
  return null
}

/** rel Type suffix that marks the embedded-workbook relationship */
const isPackageRel = (rel: Relationship): boolean =>
  rel.type.endsWith(PACKAGE_REL_SUFFIX) && !rel.target.toLowerCase().includes('.bin')

/**
 * Find a chart's embedded workbook part path (chart part rels → …/package
 * relationship → ppt/embeddings/*.xlsx). OLE2 (.bin) embeddings are skipped —
 * they are not xlsx and cannot be patched here. Returns null when absent.
 */
export function findChartEmbeddingPath(archive: PackageArchive, chartPath: string): string | null {
  for (const rel of archive.readRels(chartPath).values()) {
    if (!isPackageRel(rel)) continue
    return resolveTarget(chartPath, rel.target)
  }
  return null
}

/** next free ppt/embeddings name, PowerPoint-style Microsoft_Excel_WorksheetN.xlsx */
function nextEmbeddingName(archive: PackageArchive): string {
  let maxN = 0
  let plainTaken = false
  for (const path of archive.entries.keys()) {
    const m = /^ppt\/embeddings\/Microsoft_Excel_Worksheet(\d*)\.xlsx$/i.exec(path)
    if (!m) continue
    if (m[1] === '') plainTaken = true
    else maxN = Math.max(maxN, Number(m[1]))
  }
  return plainTaken || maxN > 0
    ? `Microsoft_Excel_Worksheet${maxN + 1}.xlsx`
    : 'Microsoft_Excel_Worksheet.xlsx'
}

/** inject <c:externalData> as the last chartSpace child (schema position) */
function ensureExternalData(archive: PackageArchive, chartPath: string, rId: string): void {
  const xml = archive.readText(chartPath)
  if (!xml) return
  if (/<c:externalData\b/.test(xml)) return
  const external = `<c:externalData r:id="${rId}"><c:autoUpdate val="0"/></c:externalData>`
  const at = xml.indexOf('</c:chartSpace>')
  if (at < 0) return
  archive.entries.set(chartPath, Buffer.from(xml.slice(0, at) + external + xml.slice(at), 'utf8'))
}

/**
 * Write the chart data into the chart's embedded workbook: patches the
 * existing ppt/embeddings/*.xlsx data rectangle in place when the chart
 * carries one, otherwise mints a fresh workbook and wires it up (chart part
 * rels …/package relationship + <c:externalData> + [Content_Types] xlsx
 * Default). Returns false when there is no chart part or the existing
 * workbook is unreadable (left untouched rather than destroyed).
 */
export function attachChartWorkbook(
  archive: PackageArchive,
  chartPath: string,
  table: ChartWorkbookTable,
): boolean {
  if (!archive.readText(chartPath)) return false
  const existing = findChartEmbeddingPath(archive, chartPath)
  if (existing) {
    const bytes = archive.readBytes(existing)
    if (!bytes) return false
    const patched = patchChartWorkbookBytes(bytes, table)
    if (!patched) return false
    archive.entries.set(existing, patched)
    return true
  }

  // mint a fresh embedding + wire the chart part to it
  const name = nextEmbeddingName(archive)
  archive.entries.set(`ppt/embeddings/${name}`, buildChartWorkbookBytes(table))
  const relsPath = relsPathFor(chartPath)
  const rels =
    archive.readText(relsPath) ??
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let maxRid = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rId = `rId${maxRid + 1}`
  const relXml = `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/${name}"/>`
  archive.entries.set(
    relsPath,
    Buffer.from(rels.replace('</Relationships>', `${relXml}</Relationships>`), 'utf8'),
  )
  ensureExternalData(archive, chartPath, rId)
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (ct && !/Extension="xlsx"/i.test(ct)) {
    archive.entries.set(
      ctPath,
      Buffer.from(ct.replace('</Types>', `${XLSX_DEFAULT}</Types>`), 'utf8'),
    )
  }
  return true
}

/**
 * ChartModel → workbook table: series names (positional fallback) and values;
 * scatter/bubble keep the x values in the category column (the same layout
 * buildChartSpaceXml's data references spell: Sheet1!$A$2:$A$n).
 */
export function chartWorkbookTableFromModel(model: ChartModel): ChartWorkbookTable {
  let categories = model.categories
  // bubble rides the scatter pipeline (kind 'scatter', y in values): the x
  // values live on the series, the category column mirrors them
  if (model.kind === 'scatter') {
    const xs = model.series.find((s) => s.xValues?.length)?.xValues
    if (xs?.length) categories = xs.map((x) => (x == null ? '' : String(x)))
  }
  return {
    categories,
    series: model.series.map((s, i) => ({
      name: s.name ?? `Series${i + 1}`,
      values: s.values,
    })),
  }
}

/**
 * Sync a chart part's embedded workbook with the data cached in the chart
 * XML itself (the part is the render source of truth; the workbook is what
 * "Edit Data" shows). Best-effort: a missing/unparseable part or workbook
 * leaves the package unchanged.
 */
export function syncChartWorkbookForChart(archive: PackageArchive, chartPath: string): boolean {
  const xml = archive.readText(chartPath)
  if (!xml) return false
  const model = parseChartXml(xml)
  if (!model) return false
  return attachChartWorkbook(archive, chartPath, chartWorkbookTableFromModel(model))
}

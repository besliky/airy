/// CSV → minimal xlsx conversion for the open-file path. Values only: cells
/// that look like plain numbers become numeric, everything else stays text
/// (leading zeros survive). The converted file is a fresh workbook.

import { decodeTextBytes } from '@airy-office/file-parse/text'

import { zipFiles } from './xlsx-minizip'

const DELIMITERS = [',', ';', '\t'] as const

/**
 * Decodes CSV bytes: BOM, then strict UTF-8, then the legacy charsets Excel
 * writes (windows-1251 included), scored script-aware in @airy-office/file-parse.
 * `preferred` (from the UI language) breaks the exact-score ties those
 * charsets produce — GBK and Shift_JIS both decode the same bytes to
 * plausible-looking but different CJK. A BOM is decoded to a U+FEFF character
 * like any text read, then dropped: the grid never wants it in a cell.
 */
export function decodeCsvBuffer(bytes: Uint8Array, preferred?: string): string {
  const text = decodeTextBytes(bytes, preferred)
  return text.startsWith('\uFEFF') ? text.slice(1) : text
}

/// Counts delimiter occurrences outside quotes over the first lines and
/// picks the most frequent one; ties favor the comma.
export function sniffDelimiter(text: string): string {
  const sample = text
    .slice(0, 64 * 1024)
    .split(/\r?\n/)
    .slice(0, 20)
  const counts = new Map<string, number>(DELIMITERS.map((d) => [d, 0]))
  for (const line of sample) {
    let quoted = false
    for (const character of line) {
      if (character === '"') quoted = !quoted
      else if (!quoted && counts.has(character)) {
        counts.set(character, (counts.get(character) ?? 0) + 1)
      }
    }
  }
  let best: string = DELIMITERS[0]
  let bestCount = -1
  for (const delimiter of DELIMITERS) {
    const count = counts.get(delimiter) ?? 0
    if (count > bestCount) {
      best = delimiter
      bestCount = count
    }
  }
  return best
}

export function parseCsv(input: string, delimiter = sniffDelimiter(input)): string[][] {
  const text = input.startsWith('\uFEFF') ? input.slice(1) : input
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 1
        } else {
          quoted = false
        }
      } else {
        field += character
      }
      continue
    }
    if (character === '"' && field === '') {
      quoted = true
    } else if (character === delimiter) {
      row.push(field)
      field = ''
    } else if (character === '\n' || character === '\r') {
      if (character === '\r' && text[index + 1] === '\n') index += 1
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else {
      field += character
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  // A trailing newline produces one empty row — drop it.
  while (rows.length > 0 && rows[rows.length - 1]?.every((cell) => cell === '')) rows.pop()
  return rows
}

/// Plain decimal numbers only; leading zeros ("007") stay text so codes and
/// phone numbers survive the import.
export function isNumericCell(value: string): boolean {
  if (!/^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?$/.test(value)) return false
  return Number.isFinite(Number(value))
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/**
 * Cell-text escape for the worksheet part: XML metacharacters plus the
 * OOXML `_xHHHH_` escape set handled by encodeXlsxEscapes (a literal `_x…_`
 * re-escapes its underscore; freshly emitted escapes are never rescanned).
 * Fusing the passes matters at import scale — 500k rows x 8 columns scan
 * this four million times.
 */
function escapeXmlCellText(value: string): string {
  return value.replace(
    // eslint-disable-next-line no-control-regex -- the control range is the thing being escaped
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF\r]|_(?=x[0-9A-Fa-f]{4}_)|[&<>"]/g,
    (character) => {
      switch (character) {
        case '&':
          return '&amp;'
        case '<':
          return '&lt;'
        case '>':
          return '&gt;'
        case '"':
          return '&quot;'
        case '_':
          return '_x005F_'
        default:
          return `_x${character.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`
      }
    },
  )
}

function columnLabel(column: number): string {
  let label = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    label = String.fromCharCode(65 + (remaining % 26)) + label
    remaining = Math.floor(remaining / 26)
  }
  return label
}

export function buildWorksheetXml(rows: readonly (readonly string[])[]): string {
  const lines: string[] = []
  let maxColumns = 1
  rows.forEach((row, rowIndex) => {
    const cells: string[] = []
    row.forEach((value, columnIndex) => {
      if (value === '') return
      maxColumns = Math.max(maxColumns, columnIndex + 1)
      const reference = `${columnLabel(columnIndex)}${rowIndex + 1}`
      cells.push(
        isNumericCell(value)
          ? `<c r="${reference}"><v>${value}</v></c>`
          : `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeXmlCellText(value)}</t></is></c>`,
      )
    })
    if (cells.length > 0) lines.push(`<row r="${rowIndex + 1}">${cells.join('')}</row>`)
  })
  const dimension = `A1:${columnLabel(maxColumns - 1)}${Math.max(rows.length, 1)}`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="${dimension}"/><sheetData>${lines.join('')}</sheetData></worksheet>`
  )
}

export async function csvToXlsxBuffer(csvText: string, sheetName = 'Sheet1'): Promise<Uint8Array> {
  const rows = parseCsv(csvText)
  if (rows.length === 0) throw new Error('The CSV file has no data rows.')
  return xlsxBufferFromRows(rows, sheetName)
}

/**
 * xlsx from the app's OWN comma-serialized sheet grid (AI create_document):
 * unlike the import path above, the delimiter is fixed to comma — cell text
 * may legitimately hold more semicolons/tabs than commas (csvField quotes
 * neither), and sniffing would then split the wrong columns — and an
 * all-empty grid becomes a valid blank workbook instead of an import error.
 */
export function sheetCsvToXlsxBuffer(csvText: string, sheetName = 'Sheet1'): Promise<Uint8Array> {
  return xlsxBufferFromRows(parseCsv(csvText, ','), sheetName)
}

/** minimal empty workbook: the backing file for a "new blank spreadsheet" tab */
export function blankXlsxBuffer(sheetName = 'Sheet1'): Promise<Uint8Array> {
  return xlsxBufferFromRows([], sheetName)
}

// The part data must build in the sandboxed renderer too (this module rides
// the lazy csv-import chunk): TextEncoder instead of Buffer.from.
const textEncoder = new TextEncoder()

async function xlsxBufferFromRows(
  rows: readonly (readonly string[])[],
  sheetName: string,
): Promise<Uint8Array> {
  // Hand-rolled zip container on Node (see xlsx-minizip): JSZip's pure-JS
  // deflate dominated the 500k-row CSV import profile, node:zlib does the
  // same DEFLATE an order of magnitude faster, and the package layout here
  // is fixed and tiny. The sandboxed renderer (no node builtins) falls back
  // to JSZip inside xlsx-minizip.
  return zipFiles([
    {
      name: '[Content_Types].xml',
      data: textEncoder.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
          '</Types>',
      ),
    },
    {
      name: '_rels/.rels',
      data: textEncoder.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
          '</Relationships>',
      ),
    },
    {
      name: 'xl/workbook.xml',
      data: textEncoder.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          `<sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
      ),
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: textEncoder.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
          '</Relationships>',
      ),
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      data: textEncoder.encode(buildWorksheetXml(rows)),
    },
  ])
}

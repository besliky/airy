/**
 * Table row-step probe (PAR-109 phase C): a two-section docx (grid 312 then
 * no grid) with small CJK tables varying line rule x spacing x snap x borders,
 * plus exact-line rows that pin the border/cell chrome. Render with the local
 * LibreOffice and read row pitches from pdftotext -bbox to decompose the LO
 * table row step into line + spacing + chrome. Measured on LO 26.2.5.2 (pure
 * CJK 12pt cells, pitch 312): single 240 = 31.2pt (2 cells), 276 auto =
 * 33.25pt = snap + 0.15 x 13.67pt, 312 auto = 35.3pt (2 x the increment),
 * after=160 at face value, borders 0.5pt, cell padding 0; doc 15's recorded
 * 24.2 rows read the same 41.75pt step.
 *
 * Run: npx tsx apps/docs/scripts/generate-table-row-probe.ts
 * Output: apps/docs/tests/pagination-corpus/probe/table-row-probe.docx
 */
import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const __dirname = dirname(fileURLToPath(import.meta.url))
const OUT_PATH = join(__dirname, '../tests/pagination-corpus/probe/table-row-probe.docx')
mkdirSync(dirname(OUT_PATH), { recursive: true })

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const STYLES_XML =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri"/>' +
  '<w:sz w:val="24"/><w:szCs w:val="24"/>' +
  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
  '<w:spacing w:after="160" w:line="276" w:lineRule="auto"/>' +
  '</w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '</w:styles>'

const sectPr = (grid: boolean) =>
  '<w:sectPr>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/>' +
  (grid ? '<w:docGrid w:type="lines" w:linePitch="312"/>' : '') +
  '</w:sectPr>'

const CN = ['一', '二', '三', '四', '五', '六', '七', '八', '九']

interface Case {
  id: string
  lineTwips: number
  rule: 'auto' | 'exact'
  after: number
  borders: boolean
  snapOff?: boolean
}

const border = (n: string) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
const withBorders =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('') +
  '</w:tblBorders>'

function table(c: Case, rows: number): string {
  const colWidth = Math.floor(8000 / 3)
  const gridCols = Array.from({ length: 3 }, () => `<w:gridCol w:w="${colWidth}"/>`).join('')
  let xml =
    `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>${c.borders ? withBorders : ''}</w:tblPr>` +
    `<w:tblGrid>${gridCols}</w:tblGrid>`
  for (let r = 0; r < rows; r++) {
    xml += '<w:tr>'
    for (let col = 0; col < 3; col++) {
      const snap = c.snapOff ? '<w:snapToGrid w:val="0"/>' : ''
      const cellText = `第${CN[r] ?? '零'}行第${CN[col] ?? '零'}列`
      xml +=
        `<w:tc><w:tcPr><w:tcW w:w="${colWidth}" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:pPr>${snap}<w:spacing w:after="${c.after}" w:line="${c.lineTwips}" w:lineRule="${c.rule}"/></w:pPr>` +
        `<w:r><w:t>${cellText}</w:t></w:r></w:p></w:tc>`
    }
    xml += '</w:tr>'
  }
  xml += '</w:tbl>'
  return xml
}

function marker(text: string): string {
  return `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p>`
}

// grid-section cases: A default (276/160), B single, C after0, D snapOff,
// E no borders, X exact 600 (30pt) after0, Y exact 600 after 160
const GRID_CASES: Case[] = [
  { id: 'A', lineTwips: 276, rule: 'auto', after: 160, borders: true },
  { id: 'B', lineTwips: 240, rule: 'auto', after: 160, borders: true },
  { id: 'C', lineTwips: 276, rule: 'auto', after: 0, borders: true },
  { id: 'D', lineTwips: 276, rule: 'auto', after: 160, borders: true, snapOff: true },
  { id: 'E', lineTwips: 276, rule: 'auto', after: 160, borders: false },
  { id: 'X', lineTwips: 600, rule: 'exact', after: 0, borders: true },
  { id: 'Y', lineTwips: 600, rule: 'exact', after: 160, borders: true },
  { id: 'Z', lineTwips: 312, rule: 'auto', after: 160, borders: true },
  { id: 'W', lineTwips: 240, rule: 'auto', after: 0, borders: true },
]

// no-grid cases: same matrix minus grid-only bits
const NOGRID_CASES: Case[] = [
  { id: 'F', lineTwips: 276, rule: 'auto', after: 160, borders: true },
  { id: 'G', lineTwips: 240, rule: 'auto', after: 160, borders: true },
  { id: 'H', lineTwips: 276, rule: 'auto', after: 0, borders: true },
  { id: 'J', lineTwips: 600, rule: 'exact', after: 0, borders: true },
  { id: 'K', lineTwips: 600, rule: 'exact', after: 160, borders: true },
  { id: 'L', lineTwips: 240, rule: 'auto', after: 0, borders: true },
]

let gridBody = ''
for (const c of GRID_CASES) {
  gridBody += marker(`case ${c.id}`) + table(c, 6) + '<w:p/>'
}
// body-paragraph references under grid: 2-line CJK paras, after 160 / after 0
gridBody += marker('case P body 276 a160')
for (let i = 0; i < 6; i++) {
  gridBody +=
    `<w:p><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:t>${CN[i]}在中经济社会发展的重要历史时期人工智能技术应用正在深刻改变各行各业的生产方式和商业模式数字化转型已成为关键</w:t></w:r></w:p>`
}
gridBody += marker('case Q body 276 a0')
for (let i = 0; i < 6; i++) {
  gridBody +=
    `<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:t>${CN[i]}在中经济社会发展的重要历史时期人工智能技术应用正在深刻改变各行各业的生产方式和商业模式数字化转型已成为关键</w:t></w:r></w:p>`
}
// close the grid section here: everything above is grid, everything below is not
let noGridBody = `<w:p><w:pPr>${sectPr(true)}</w:pPr></w:p>`
for (const c of NOGRID_CASES) {
  noGridBody += marker(`case ${c.id}`) + table(c, 6) + '<w:p/>'
}

const documentXml =
  XML_DECL +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${gridBody + noGridBody + sectPr(false)}</w:body></w:document>`

async function main() {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  zip.file('word/document.xml', documentXml)
  zip.file('word/styles.xml', STYLES_XML)
  const buf = await zip.generateAsync({ type: 'nodebuffer' })
  writeFileSync(OUT_PATH, buf)
  console.log('written', OUT_PATH)
}

void main()

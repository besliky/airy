/**
 * Grid-snap probe matrix (PAR-109 spec §1.4): one docx per pitch with
 * SimSun-class 12/13/14/16/18pt paragraphs under single (240) and 1.15 (276) auto
 * line rules. Render on the target engine (real Word for Mac / LibreOffice)
 * and read line pitches from the PDF to decide the grid snap base:
 * em-box vs natural-height snapping (phase A) and the profile arithmetic
 * (phase B). Output: tests/pagination-corpus/probe/grid-probe-<pitch>.docx
 *
 * Run: npx tsx apps/docs/scripts/generate-grid-probe.ts
 */

import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const __dirname = dirname(fileURLToPath(import.meta.url))
const OUT_DIR = join(__dirname, '../tests/pagination-corpus/probe')

mkdirSync(OUT_DIR, { recursive: true })

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const STYLES_XML =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri"/>' +
  '<w:sz w:val="24"/><w:szCs w:val="24"/>' +
  '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
  '<w:spacing w:after="160"/>' +
  '</w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '</w:styles>'

/** A4 + margins matching the corpus grid docs (content width 553.7px @96dpi) */
const sectPr = (pitch: number) =>
  '<w:sectPr>' +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/>' +
  `<w:docGrid w:type="lines" w:linePitch="${pitch}"/>` +
  '</w:sectPr>'

const ZH =
  '在中经济社会发展的重要历史时期人工智能技术广泛应用正在深刻改变各行各业生产方式和商业模式数字化转型已成为企业保持竞争力的关键战略各类组织机构纷纷加快推进信息化建设步伐'

interface Case {
  sizePt: number
  lineTwips: number
}

const CASES: Case[] = []
for (const sizePt of [12, 13, 14, 16, 18]) {
  for (const lineTwips of [240, 276]) {
    CASES.push({ sizePt, lineTwips })
  }
}

function casePara(c: Case, index: number): string {
  const halfPoints = c.sizePt * 2
  const rule = c.lineTwips === 240 ? 'auto' : 'auto'
  // a short marker paragraph then three probe paragraphs at the case size/rule
  const marker =
    `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>` +
    `<w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>case ${index}: ${c.sizePt}pt line ${c.lineTwips}</w:t></w:r></w:p>`
  const probe = (n: number) =>
    `<w:p><w:pPr><w:spacing w:after="0" w:line="${c.lineTwips}" w:lineRule="${rule}"/></w:pPr>` +
    `<w:r><w:rPr><w:sz w:val="${halfPoints}"/></w:rPr><w:t>${ZH.slice(0, 34 * n)}</w:t></w:r></w:p>`
  return marker + probe(2) + probe(2) + probe(2) + probe(2)
}

async function buildDocx(pitch: number, bodyXml: string): Promise<Uint8Array> {
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
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}${sectPr(pitch)}</w:body></w:document>`,
  )
  zip.file('word/styles.xml', STYLES_XML)
  return zip.generateAsync({ type: 'uint8array' })
}

async function main() {
  for (const pitch of [312, 360]) {
    const body = CASES.map((c, i) => casePara(c, i + 1)).join('')
    const bytes = await buildDocx(pitch, body)
    const out = join(OUT_DIR, `grid-probe-${pitch}.docx`)
    writeFileSync(out, bytes)
    console.log(`wrote ${out}`)
  }
}

void main()

// Fixture builder for the docx tests: a minimal, valid .docx assembled with
// jszip in-process (no network, no committed binaries). Mirrors the shape of
// packages/docx-engine/tests/helpers/build-docx.ts, kept local so the MCP
// package's tests stay self-contained.
import JSZip from 'jszip'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const STYLES_XML =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>'

const NUMBERING_XML =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '</w:numbering>'

const BODY_XML = [
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Quarterly Report</w:t></w:r></w:p>',
  '<w:p><w:r><w:t xml:space="preserve">Revenue grew by </w:t></w:r>' +
    '<w:r><w:rPr><w:b/></w:rPr><w:t>12 percent</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve"> year over year.</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First bullet item</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second bullet item</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Numbered step</w:t></w:r></w:p>',
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Sales</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>East</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>4200</w:t></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl>',
  '<w:p><w:r><w:t>End of report.</w:t></w:r></w:p>',
].join('')

// jszip stamps every entry with the CURRENT date when no date option is given
// (DOS timestamp, 2-second granularity), so two builds that straddle a
// 2-second boundary produce different header bytes. Tests byte-compare a
// freshly built fixture against bytes written at setup time ("original on
// disk unchanged"), which made that assertion time-dependent and flaky on CI
// (TEST-721). The date option only pins entries added EXPLICITLY: jszip also
// auto-creates a directory entry for every parent of a nested path
// (fileAdd -> folderAdd) and that internal call takes no date, so implicit
// dirs (here: _rels/, word/ AND word/_rels/) would silently keep the build
// time. addPinned() creates every ancestor explicitly with the fixed date,
// which makes the whole builder byte-deterministic however it grows.
const FIXED_ZIP_DATE = { date: new Date(Date.UTC(2024, 1, 2, 3, 4, 6)) }

function addPinned(zip: JSZip, name: string, content: string): void {
  const segments = name.split('/')
  segments.pop()
  let prefix = ''
  for (const segment of segments) {
    prefix += `${segment}/`
    if (zip.files[prefix] === undefined) {
      zip.file(prefix, null, { dir: true, ...FIXED_ZIP_DATE })
    }
  }
  zip.file(name, content, FIXED_ZIP_DATE)
}

/** the representative test document: heading, styled paragraph, lists, table, tail */
export async function buildFixtureDocx(): Promise<Uint8Array> {
  return buildBodyDocx(BODY_XML)
}

/**
 * Arbitrary-body fixture with the standard parts (styles + numbering from the
 * representative document): lets a test spell exact table/paragraph XML for
 * op-level edge cases without a dedicated builder per shape.
 */
export async function buildBodyDocx(bodyXml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  addPinned(
    zip,
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>',
  )
  addPinned(
    zip,
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  addPinned(
    zip,
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      '</Relationships>',
  )
  addPinned(zip, 'word/styles.xml', STYLES_XML)
  addPinned(zip, 'word/numbering.xml', NUMBERING_XML)
  addPinned(
    zip,
    'word/document.xml',
    `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/**
 * Synthetic overview-flood document: `blockCount` plain paragraphs ("Block N
 * …" text), for the read-budget regression that needs a document whose
 * overview alone dwarfs the 30k answer budget (BUG-1684 mirrors the audited
 * 42k-block corpus). Each paragraph is index-addressable and keeps a stable,
 * greppable text.
 */
export async function buildManyBlocksDocx(blockCount: number): Promise<Uint8Array> {
  const paragraphs: string[] = []
  for (let i = 0; i < blockCount; i++) {
    paragraphs.push(`<w:p><w:r><w:t>Block ${i} lorem ipsum dolor</w:t></w:r></w:p>`)
  }
  const zip = new JSZip()
  addPinned(
    zip,
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  addPinned(
    zip,
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  addPinned(
    zip,
    'word/document.xml',
    `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs.join('')}` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

// BUG-1502 fixture: the list numbering lives on the ListBullet/ListNumber
// styles (w:numPr inside the style's w:pPr), and the body paragraphs carry
// only the pStyle — no direct w:numPr. This is how Word writes documents
// where the author picked "List Bullet" from the style gallery.
const STYLE_LIST_STYLES_XML =
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style>' +
  '</w:styles>'

const STYLE_LIST_BODY_XML = [
  '<w:p><w:r><w:t>Intro paragraph.</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Styled bullet one</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Styled bullet two</w:t></w:r></w:p>',
  '<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>Styled numbered step</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>Outro paragraph.</w:t></w:r></w:p>',
].join('')

/** variant with style-driven lists (numbering via pStyle, no direct numPr) */
export async function buildStyleListDocx(): Promise<Uint8Array> {
  const zip = new JSZip()
  addPinned(
    zip,
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>',
  )
  addPinned(
    zip,
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  addPinned(
    zip,
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      '</Relationships>',
  )
  addPinned(zip, 'word/styles.xml', STYLE_LIST_STYLES_XML)
  addPinned(zip, 'word/numbering.xml', NUMBERING_XML)
  addPinned(
    zip,
    'word/document.xml',
    `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${STYLE_LIST_BODY_XML}` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

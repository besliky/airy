// Fixture builders for the legacy/ODF tests. The .odt is a minimal valid
// OpenDocument zip assembled in-process (mimetype first and stored, per the
// ODF packaging spec) — LibreOffice opens it through the writer8 path.
// Binary .xls/.doc fixtures cannot be generated without third-party writers;
// .doc reuses the committed sample from @genoffice/file-parse, .xls is only
// exercised when a real fixture is supplied (see the skipIf in the tests).
import JSZip from 'jszip'

const CONTENT_XML =
  '<?xml version="1.0" encoding="UTF-8"?>\n' +
  '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'office:version="1.2">' +
  '<office:body><office:text>' +
  '<text:h text:outline-level="1">ODF Source Document</text:h>' +
  '<text:p>Original ODT paragraph one.</text:p>' +
  '<text:p>Original ODT paragraph two.</text:p>' +
  '</office:text></office:body></office:document-content>'

const MANIFEST_XML =
  '<?xml version="1.0" encoding="UTF-8"?>\n' +
  '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" ' +
  'manifest:version="1.2">' +
  '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" ' +
  'manifest:media-type="application/vnd.oasis.opendocument.text"/>' +
  '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
  '</manifest:manifest>'

/** Minimal valid .odt (one heading + two paragraphs). */
export async function buildOdtFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  // mimetype must be the first entry, uncompressed
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', MANIFEST_XML)
  zip.file('content.xml', CONTENT_XML)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

const SPREADSHEET_CONTENT_XML =
  '<?xml version="1.0" encoding="UTF-8"?>\n' +
  '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'office:version="1.2">' +
  '<office:body><office:spreadsheet>' +
  '<table:table table:name="Sheet1">' +
  '<table:table-row>' +
  '<table:table-cell office:value-type="string"><text:p>Alpha</text:p></table:table-cell>' +
  '<table:table-cell office:value-type="string"><text:p>Beta</text:p></table:table-cell>' +
  '</table:table-row>' +
  '<table:table-row>' +
  '<table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>' +
  '<table:table-cell office:value-type="float" office:value="2"><text:p>2</text:p></table:table-cell>' +
  '</table:table-row>' +
  '</table:table>' +
  '</office:spreadsheet></office:body></office:document-content>'

const SPREADSHEET_MANIFEST_XML =
  '<?xml version="1.0" encoding="UTF-8"?>\n' +
  '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" ' +
  'manifest:version="1.2">' +
  '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" ' +
  'manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>' +
  '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
  '</manifest:manifest>'

/** Minimal valid .ods (one 2x2 sheet: strings + floats) readable by calamine. */
export async function buildOdsFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' })
  zip.file('META-INF/manifest.xml', SPREADSHEET_MANIFEST_XML)
  zip.file('content.xml', SPREADSHEET_CONTENT_XML)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/**
 * Native document parts: real Word tables, list numbering, content controls,
 * text runs, fonts, and document metadata — the converter's editable core.
 * One Chrome is shared by every case in this file (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { EXTRACTOR_SOURCE, mapFont } from '../src'
import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

test('assembles a self-contained browser extractor function', () => {
  assert.match(EXTRACTOR_SOURCE, /^function extractIR\(\) \{/)
  assert.match(EXTRACTOR_SOURCE, /function markForScreenshot/)
  assert.match(EXTRACTOR_SOURCE, /function extractTable/)
  assert.match(EXTRACTOR_SOURCE, /function processElement/)
  assert.match(EXTRACTOR_SOURCE, /__html2docxPages\.build/)
})

test('preserves rowspan, numbering, controls, headers, media, and page size', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html>
    <html><head><style>
      @page { size: letter landscape; margin: 0.5in; }
      body { max-width: 900px; margin: 0 auto; font-family: Arial; }
      .row { display:flex; gap:24px; }
      .row > div { width:180px; border:1px solid #999; padding:8px; }
      .shadow { width:240px; padding:16px; box-shadow:0 4px 12px rgba(0,0,0,.3); }
    </style></head><body>
      <header data-docx-header><p>Repeated header</p></header>
      <table><tbody>
        <tr><th rowspan="2">Merged</th><th>A</th></tr>
        <tr><td>B</td></tr>
      </tbody></table>
      <ol style="list-style-type:lower-alpha"><li>Alpha</li><li>Beta</li></ol>
      <input name="customer" value="Alex">
      <input type="checkbox" name="approved" checked>
      <select name="status"><option>Draft</option><option selected>Final</option></select>
      <div class="row"><div>Left</div><div>Right</div></div>
      <div class="shadow">Shadow content</div>
      <iframe title="Frame preview" srcdoc="<p>Embedded frame</p>" style="width:240px;height:100px"></iframe>
    </body></html>`,
    'features',
  )
  const documentXml = await xml('word/document.xml')
  const numberingXml = await xml('word/numbering.xml')
  const headerXml = await xml('word/header1.xml')
  assert.match(documentXml, /w:vMerge w:val="restart"/)
  assert.match(documentXml, /w:vMerge w:val="continue"/)
  assert.match(documentXml, /w:sdt/)
  assert.match(documentXml, /w14:checkbox/)
  assert.match(documentXml, /w:dropDownList/)
  assert.match(documentXml, /w:pgSz w:w="15840" w:h="12240"/)
  assert.match(numberingXml, /w:numFmt w:val="lowerLetter"/)
  assert.match(headerXml, /Repeated header/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'shadow and iframe should use screenshot media',
  )
})

test('preserves CSS display-table rows as one native table', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .competencies { display:table; width:700px; }
      .row { display:table-row; }
      .code, .description { display:table-cell; padding:12px; }
      .code { width:120px; color:white; background:#1a5276; }
      .description { background:#dceef8; }
    </style></head><body>
      <div class="competencies">
        <div class="row"><div class="code">KD 3.2</div><div class="description">Analyze quantities</div></div>
        <div class="row"><div class="code">KD 4.2</div><div class="description">Present measurements</div></div>
      </div>
    </body></html>`,
    'css-display-table',
  )
  const documentXml = await xml('word/document.xml')
  assert.equal((documentXml.match(/<w:tbl>/g) || []).length, 1)
  assert.equal((documentXml.match(/<w:tr>/g) || []).length, 2)
  assert.equal((documentXml.match(/<w:tc>/g) || []).length, 4)
  assert.match(documentXml, /w:fill="1A5276"/)
  assert.match(documentXml, /w:fill="DCEEF8"/)
})

test('removes default Word paragraph spacing inside compact table cells', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      table { border-collapse:collapse; width:600px; font-size:12px; }
      td { padding:7px 12px; border:1px solid #d5d8dc; }
    </style></head><body>
      <table><tbody>
        <tr><td>Panjang</td><td>Keindahan</td></tr>
        <tr><td>Massa</td><td>Kejujuran</td></tr>
      </tbody></table>
    </body></html>`,
    'compact-table-cell-spacing',
  )
  const documentXml = await xml('word/document.xml')
  assert.ok((documentXml.match(/<w:spacing w:after="0" w:before="0"\/>/g) || []).length >= 4)
  assert.doesNotMatch(documentXml, /<w:trHeight/)
})

test('removes paragraph margins inside padded table cells', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body><table><tr><td style="padding:8px">
      <p style="margin:12px 0">Compact row</p>
    </td></tr></table></body></html>`,
    'nested-cell-paragraph-spacing',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:spacing w:after="0" w:before="0"/)
})

test('renders closed details as a static summary without hidden content', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><body>
      <details style="border:1px solid #86efac;background:#f0fdf4;padding:8px 14px">
        <summary>Show answer</summary>
        <p>Hidden answer text</p>
      </details>
    </body></html>`,
    'closed-details',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Show answer/)
  assert.doesNotMatch(documentXml, /Hidden answer text/)
  assert.equal(
    ir.some((node) => node.type === 'card'),
    false,
  )
  assert.ok(ir.some((node) => node.type === 'para' && node.style?.shading === 'F0FDF4'))
})

test('splits emoji icons into font-unset runs for Word fallback', async () => {
  const { xml } = await convertHtml(
    '<!doctype html><html><body><h2>🏆 Performance 🟢</h2></body></html>',
    'emoji-icons',
  )
  const documentXml = await xml('word/document.xml')
  // Each emoji lands in its own run whose rPr carries no rFonts, so Word's
  // built-in color-emoji fallback applies. Forcing Segoe UI Emoji rendered
  // tofu (and flags as letters) on Mac Word.
  assert.match(documentXml, /<w:rPr>(?:(?!<\/w:rPr>|w:rFonts)[^])*?<\/w:rPr><w:t[^>]*>🏆<\/w:t>/)
  assert.match(documentXml, /<w:rPr>(?:(?!<\/w:rPr>|w:rFonts)[^])*?<\/w:rPr><w:t[^>]*>🟢<\/w:t>/)
  // The adjacent body word still gets a real font.
  assert.match(documentXml, /<w:rFonts[^>]*\/>[^]*?<w:t[^>]*> Performance <\/w:t>/)
})

test('maps rounded display fonts before generic cursive fallbacks', () => {
  assert.equal(mapFont('"Fredoka One", cursive'), 'Arial Rounded MT Bold')
  assert.equal(mapFont('"Pacifico", cursive'), 'Segoe Script')
  assert.equal(mapFont('"Open Sans", sans-serif'), 'Arial')
})

test('keeps authored newlines in pre-wrap text blocks', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <div style="white-space:pre-wrap;">Hi [First Name],

I am reaching out about the launch.
Second line of the same paragraph.

Closing line.</div>
    </body></html>`,
    'pre-wrap-newlines',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Hi \[First Name\],/)
  const breaks = (documentXml.match(/<w:br\/>/g) || []).length
  assert.ok(breaks >= 4, `authored newlines should become breaks, got ${breaks}`)
})

test('maps document meta, html lang, and footer page numbers to native Word parts', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html lang="ja"><head>
      <title>四半期業績レポート</title>
      <meta name="author" content="山田太郎"><!-- public-hygiene: fixture -->
      <meta name="description" content="第2四半期の業績分析">
      <style>footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; }</style>
    </head><body>
      <h1>業績サマリー</h1>
      <p>売上高は前年同期比 12% 増となりました。ひらがなとカタカナのテキストです。</p>
      <footer>Page 1 of 3 — 社外秘</footer><!-- public-hygiene: fixture -->
    </body></html>`,
    'doc-meta-lang-pagenum',
  )
  const core = await xml('docProps/core.xml')
  assert.match(core, /<dc:title>四半期業績レポート<\/dc:title>/)
  assert.match(core, /<dc:creator>山田太郎<\/dc:creator>/) // public-hygiene: fixture
  const styles = await xml('word/styles.xml')
  assert.match(styles, /w:eastAsia="ja-JP"/)
  const footer = await xml('word/footer1.xml')
  assert.match(footer, />PAGE</)
  assert.match(footer, />NUMPAGES</)
  assert.match(footer, /社外秘/) // public-hygiene: fixture
})

test('keeps broken image alt text as an editable fallback', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <div><img src="file:///definitely-missing-image.png" alt="Rocket"><strong>FAUGET CO.</strong></div>
    </body></html>`,
    'broken-image-alt',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Rocket/)
  assert.match(documentXml, /FAUGET CO\./)
})

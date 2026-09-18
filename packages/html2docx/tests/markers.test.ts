/**
 * List markers, chips, badges, and icon glyphs: pseudo-element counters,
 * literal bullets, highlighted inline chips, and small rasterized marks.
 * One Chrome is shared by every case in this file (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

test('preserves pseudo counter order and literal pseudo bullets', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      ol, ul { list-style:none; }
      ol { counter-reset:step; }
      ol li { counter-increment:step; }
      ol li::before { content:"STEP " counter(step); text-transform:uppercase; }
      ul li::before { content:"■"; color:#6b2c2c; }
    </style></head><body>
      <ol><li><div>Review request</div><p>Verify details</p></li></ol>
      <ul><li>Square marker item</li></ul>
    </body></html>`,
    'pseudo-list-markers',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /STEP 1/)
  assert.doesNotMatch(documentXml, /1STEP/)
  assert.match(documentXml, /■/)
})

test('preserves short literal markers on list items', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      ol { list-style:none; padding:0; }
      li { position:relative; padding-left:36px; }
      li::before { position:absolute; left:8px; color:#1e40af; }
      li:nth-child(1)::before { content:"① "; }
      li:nth-child(2)::before { content:"② "; }
    </style></head><body>
      <ol><li>First choice</li><li>Second choice</li></ol>
    </body></html>`,
    'literal-list-markers',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /①/)
  assert.match(documentXml, /②/)
  assert.match(documentXml, /w:ascii="Arial Unicode MS"/)
})

test('preserves pseudo symbols in multicolumn lists and labels', async () => {
  const { zip, xml, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .ribbon { display:inline-block; background:#435638; color:white; padding:4px 12px; }
      .ribbon::before { content:"✓ "; color:#d8c48d; }
      ul { columns:2; width:600px; list-style:none; padding:0; }
      li { position:relative; padding-left:24px; }
      li::before { content:"☑"; position:absolute; left:0; color:#4a5e3a; }
    </style></head><body>
      <div class="ribbon">All Included</div>
      <ul><li>One</li><li>Two</li><li>Three</li><li>Four</li></ul>
    </body></html>`,
    'multicolumn-pseudo-markers',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /✓/)
  assert.match(screenshotText, /One/)
  assert.match(screenshotText, /Four/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 4,
    'the four exact pseudo-marker rows should be screenshots',
  )
})

test('preserves colored literal bullets as editable runs', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><body>
      <div style="color:#e74c3c">• Colored bullet item</div>
    </body></html>`,
    'colored-literal-bullet',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:color w:val="E74C3C"\/>/)
  assert.match(documentXml, /• Colored bullet item/)
  assert.doesNotMatch(documentXml, /<w:numPr>/)
})

test('preserves horizontal padding on highlighted inline chips', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .dim {
        display:inline-block;
        background:#1a5276;
        color:white;
        font-weight:700;
        padding:1px 7px;
      }
    </style></head><body><p>Dimension <span class="dim">[L]</span></p></body></html>`,
    'highlighted-inline-padding',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /\[L\]/)
  assert.ok((documentXml.match(/w:fill="1A5276"/g) || []).length >= 3)
  assert.ok((documentXml.match(/>\u00a0+</g) || []).length >= 2)
})

test('screenshots translucent number badges on colored headers', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .head {
        background:#2563eb;
        color:#fff;
        padding:18px 20px;
        text-align:center;
        width:220px;
        display:flex;
        flex-direction:column;
        align-items:center;
        gap:8px;
      }
      .num {
        display:inline-flex;
        width:30px;height:30px;
        align-items:center;justify-content:center;
        border-radius:999px;
        background:rgba(255,255,255,.2);
        color:#fff;
        font-weight:700;
      }
      .tag-pill {
        display:inline-block;
        background:#f5f5f5;
        border:1px solid #ccc;
        padding:2px 8px;
        border-radius:12px;
        margin:2px 4px;
      }
    </style></head><body>
      <div class="head">
        <div class="num">1</div>
        <div>理由がわかる</div>
      </div>
      <div>
        <span class="tag-pill">#休養</span><!-- public-hygiene: fixture -->
        <span class="tag-pill">#疲労回復</span><!-- public-hygiene: fixture -->
        <span class="tag-pill">#コンディショニング</span>
      </div>
    </body></html>`,
    'translucent-badge',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /理由/) // public-hygiene: fixture
  // Must not become a white color-bar with white "1" (invisible badge).
  assert.doesNotMatch(documentXml, /w:fill="FFFFFF"[\s\S]{0,200}>1</)
  const media = Object.keys(zip.files).filter((name) => name.startsWith('word/media/'))
  assert.ok(media.length >= 1, 'circle badge should be captured as an image')
  // Number chip is one image; tag cloud may be a second combined screenshot.
  assert.ok(media.length <= 3, `too many rasterized chips, got ${media.length} media`)
  // Tag cloud is captured as one image (not editable text runs).
  assert.ok(
    media.length >= 2 || documentXml.includes('#休養'), // public-hygiene: fixture
    'tag pills should appear as text or as a combined screenshot',
  )
})

test('keeps flex icon headings on one line with their authored gap', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .heading { display:flex; align-items:center; font-size:14px; margin-bottom:25px; }
      .icon { width:20px; height:20px; margin-right:12px; }
    </style></head><body>
      <div class="heading">
        <svg class="icon" viewBox="0 0 24 24"><path d="M2 10l10-5 10 5-10 5z"></path></svg>
        GRADUATION
      </div>
    </body></html>`,
    'inline-flex-icon-heading',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:p>[\s\S]*?<w:drawing>[\s\S]*?GRADUATION[\s\S]*?<\/w:p>/)
  assert.match(documentXml, /\u00a0/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('keeps empty checkbox boxes inside flex rows', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      ul { list-style:none; padding:0; }
      li { display:flex; align-items:center; gap:12px; }
      .box { width:22px; height:22px; border:1.5px solid #cbd5e1; border-radius:6px; }
    </style></head><body>
      <ul><li><span class="box"></span><span>Unchecked criterion</span></li></ul>
    </body></html>`,
    'flex-row-checkbox',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /w:color="CBD5E1"/)
  assert.ok((documentXml.match(/<w:tbl>/g) || []).length >= 2)
  assert.match(documentXml, /Unchecked criterion/)
})

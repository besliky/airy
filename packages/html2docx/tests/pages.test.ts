/**
 * Page-level composition: repeated page backgrounds, page decorations,
 * explicit and inferred page containers, full-page screenshot slicing, and
 * page-flow guarantees (no phantom blank pages, watermarks past page one).
 * One Chrome is shared by every case in this file (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

test('repeats page backgrounds from the header without adding a body page', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#0b0b0b url("data:image/gif;base64,R0lGODlhAQABAIAAAAUEBAAAACwAAAAAAQABAAACAkQBADs=") repeat; color:#f4f1e8; }
      .second { break-before:page; }
    </style></head><body>
      <p>Dark page one</p>
      <p class="second">Dark page two</p>
    </body></html>`,
    'repeated-background',
  )
  const documentXml = await xml('word/document.xml')
  const headerXml = await xml('word/header1.xml')
  assert.match(headerXml, /wp:anchor/)
  assert.match(headerXml, /behindDoc="1"/)
  assert.match(headerXml, /w:line="1" w:lineRule="exact"/)
  assert.match(documentXml, /w:pgMar[^>]*w:header="0"/)
  assert.doesNotMatch(documentXml, /behindDoc="1"/)
  assert.match(documentXml, /Dark page one/)
  assert.match(documentXml, /Dark page two/)
})

test('uses a plain background color for flat solid page backdrops', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#0b0b0b; color:#f4f1e8; }
      .second { break-before:page; }
    </style></head><body>
      <p>Dark page one</p>
      <p class="second">Dark page two</p>
    </body></html>`,
    'flat-color-background',
  )
  const documentXml = await xml('word/document.xml')
  // no backdrop screenshot: the resampled header image's tone visibly
  // differs from same-color shading fills on content blocks
  assert.match(documentXml, /<w:background w:color="0B0B0B"\/>/)
  assert.match(documentXml, /Dark page one/)
  assert.match(documentXml, /Dark page two/)
})

test('uses a zero-height anchor paragraph for page decorations', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#f4f4f4; }
      main, h1, p { margin:0; }
      .curve {
        position:absolute;
        top:0;
        right:0;
        width:120px;
        height:80px;
        background:linear-gradient(135deg,#4a90c2,#8fb8dd);
        border-radius:0 0 0 100%;
      }
    </style></head><body>
      <div class="curve"></div>
      <main><h1>Resume title</h1><p>First content line</p></main>
    </body></html>`,
    'floating-decoration-anchor',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /wp:anchor/)
  assert.match(documentXml, /w:line="1" w:lineRule="exact"/)
  // flat-color page backgrounds use w:background, not a header image
  assert.match(documentXml, /<w:background w:color="F4F4F4"\/>/)
  // a top:0 decoration must not be mistaken for a footer bar on short bodies
  assert.doesNotMatch(documentXml, /<wp:align>bottom<\/wp:align>/)
})

test('renders explicit multipage containers as separate page images', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      [data-docx-page] { width:600px; height:800px; overflow:hidden; position:relative; }
      .second { background:#ddeeff; }
      .floating { position:absolute; left:200px; top:300px; transform:rotate(8deg); }
    </style></head><body>
      <section data-docx-page><h1>Page one</h1></section>
      <section data-docx-page class="second"><h1 class="floating">Page two</h1></section>
    </body></html>`,
    'multipage',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /w:br w:type="page"/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'each explicit page should produce a page image',
  )
})

test('keeps semantic content between inferred page compositions', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .opener {
        width:700px;
        height:800px;
        break-before:page;
        break-after:page;
        position:relative;
      }
      .opener h1 { position:absolute; left:120px; top:300px; }
      .full-page { height:1119px; }
      .chapter { break-before:page; }
    </style></head><body>
      <section class="opener full-page"><h1>Chapter one opener</h1></section>
      <section class="chapter">
        <p>Chapter one body sentinel</p>
        <ul><li>Chapter one list sentinel</li></ul>
      </section>
      <section class="opener"><h1>Chapter two opener</h1></section>
      <section class="chapter">
        <table><tr><td>Chapter two table sentinel</td></tr></table>
      </section>
    </body></html>`,
    'mixed-page-compositions',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Chapter one body sentinel/)
  assert.match(documentXml, /Chapter one list sentinel/)
  assert.match(documentXml, /Chapter two table sentinel/)
  assert.match(documentXml, /w:pgMar[^>]*w:header="0"[^>]*w:footer="0"/)
  assert.equal(
    (documentXml.match(/w:br w:type="page"/g) || []).length,
    2,
    'a full-page image should advance naturally without creating a blank page',
  )
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'chapter openers should remain page screenshots',
  )
})

test('keeps absolute-positioned content beyond the first page', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { position:relative; width:700px; height:1800px; margin:0 auto; }
      .first { position:absolute; top:80px; left:60px; }
      .second { position:absolute; top:1320px; left:180px; transform:rotate(5deg); }
    </style></head><body>
      <h1 class="first">First page</h1>
      <h1 class="second">Second page absolute content</h1>
    </body></html>`,
    'absolute-pages',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(
    documentXml,
    /w:br w:type="page"/,
    'full-page image slices should advance naturally without inserting blank pages',
  )
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'both page clips should be embedded',
  )
})

test('splits a tall visual section into flowable screenshot slices', async () => {
  const { ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .visual {
        display:block;
        width:700px;
        height:1300px;
        background:linear-gradient(#0c1228,#0a0f1e);
        background-color:#0c1228;
      }
    </style></head><body>
      <img class="visual" src="/missing-map.png" alt="Consumer map" />
      <p>Following content</p>
    </body></html>`,
    'tall-visual-slices',
  )
  const slices = ir.filter((node) => node.type === 'image' && node.clip)
  assert.equal(slices.length, 2, JSON.stringify(ir))
  assert.equal(slices[0].height, 700)
  assert.equal(slices[1].height, 600)
  assert.match(screenshotText, /Consumer map/)
})

test('rasterizes a single full-page visual poster', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      @page { size:A4; margin:0; }
      body { margin:0; display:flex; justify-content:center; }
      .poster {
        width:794px; height:1123px; position:relative; overflow:hidden;
        background:linear-gradient(#fffef8,#f7f8ef);
      }
    </style></head><body><div class="poster"><h1>Visual poster</h1><p>Composed page</p></div></body></html>`,
    'single-visual-poster',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /Visual poster/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('rasterizes a tall painted resume grid as natural page slices', async () => {
  const { zip, xml, ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; display:flex; justify-content:center; }
      .resume {
        width:700px; height:1500px; position:relative; overflow:hidden;
        background:#f5f1ed;
      }
      .header { height:220px; text-align:center; }
      .columns { display:grid; grid-template-columns:1fr 1.8fr; gap:24px; }
      .card { min-height:900px; background:white; padding:20px; }
    </style></head><body><main class="resume">
      <header class="header"><h1>Resume name</h1></header>
      <div class="columns"><section class="card">Profile</section><section class="card">Experience</section></div>
    </main></body></html>`,
    'flowing-visual-resume',
  )
  const documentXml = await xml('word/document.xml')
  assert.equal(ir.filter((node) => node.type === 'image').length, 2)
  assert.equal(ir.filter((node) => node.type === 'pagebreak').length, 0)
  assert.doesNotMatch(documentXml, /Resume name/)
  assert.match(screenshotText, /Resume name/)
  assert.match(screenshotText, /Experience/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'each resume viewport slice should be embedded as an image',
  )
})

test('moves a page-straddling contact card whole to the next slice', async () => {
  // Regression: the naive full-page-height slice boundary cut through the
  // contact card and Word's ~26px top inset then clipped the card's last
  // row ("Greater Houston…") at the page edge. Slices must cut in the gap
  // before the card and stay under the page budget.
  const { ir, xml, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; display:flex; justify-content:center; }
      .resume {
        width:700px; position:relative; overflow:hidden;
        background:#f5f1ed; padding-bottom:60px;
      }
      .head { height:200px; display:flex; gap:16px; }
      .work { height:760px; background:white; margin:0 24px; }
      .contact { height:320px; background:white; margin:40px 24px 0; padding:20px; }
      .contact p { margin:12px 0; }
    </style></head><body><main class="resume">
      <header class="head"><h1>Contact name</h1><span>General Manager</span></header>
      <section class="work">Experience body</section>
      <section class="contact">
        <p>Date of Birth: 01/01/1990</p>
        <p>(555) 010-0199</p>
        <p>contact@example.com</p>
        <p>Greater Houston, TX — open to relocate</p>
      </section>
    </main></body></html>`,
    'contact-card-slice',
  )
  const documentXml = await xml('word/document.xml')
  const slices = ir.filter((node) => node.type === 'image' && node.pageComposition)
  assert.equal(slices.length, 2)
  const budget = Math.round(1123 * 0.96)
  for (const slice of slices) {
    assert.ok(slice.height <= budget, `slice height ${slice.height} exceeds budget ${budget}`)
  }
  // The cut must land in the gap between the work and contact cards
  // (work bottom 960, contact top 1000), not inside the contact card.
  assert.ok(
    slices[0].height > 940 && slices[0].height < 1000,
    `first slice should end in the inter-card gap, got ${slices[0].height}`,
  )
  assert.match(screenshotText, /Greater Houston, TX — open to relocate/)
  assert.doesNotMatch(documentXml, /Greater Houston/)
})

test('renders an active SPA page as full-page screenshots', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      nav { position:fixed; inset:0 0 auto; height:64px; background:#071426; }
      .page-view { display:none; }
      .page-view.active-page { display:block; min-height:2400px; }
      .hero { height:1200px; background:linear-gradient(135deg,#071426,#1f4e9a); }
      .container { width:90%; margin:auto; }
      .row { display:flex; }
      img { width:400px; height:300px; }
    </style></head><body>
      <nav>Navigation</nav>
      <section class="page-view active-page">
        <div class="hero"><div class="container row"><h1>Visual website</h1>
          <img src="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='400' height='300'%3E%3Crect width='400' height='300' fill='%2379d0da'/%3E%3C/svg%3E">
        </div></div>
      </section>
    </body></html>`,
    'spa-page',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(
    documentXml,
    /w:br w:type="page"/,
    'full-page image slices should advance naturally without inserting blank pages',
  )
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 2,
    'each website viewport slice should be embedded as an image',
  )
  assert.doesNotMatch(documentXml, /Visual website/)
})

test('keeps an explicit white body over a gray browser canvas', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      html, body { background:#f0f0f0; }
      body { width:794px; min-height:1123px; margin:20px auto; padding:80px; background:#fff; }
    </style></head><body><h1>Cover letter</h1><p>White paper content</p></body></html>`,
    'white-paper-gray-canvas',
  )
  assert.equal(
    ir.some((node) => node.type === 'pagebg'),
    false,
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /<w:background w:color="F0F0F0"\/>/)
})

test('keeps section watermark numerals past the first page', async () => {
  const section = (num, tint) => `
    <section style="background:${tint}; position:relative; padding:120px 40px; height:760px; box-sizing:border-box;">
      <div class="num">${num}</div>
      <h2>Section ${num}</h2>
      <p>Section body copy.</p>
      <div style="display:flex; gap:20px;"><div>Cell A</div><div>Cell B</div></div>
    </section>`
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#f4f0e7; }
      .num { position:absolute; top:24px; left:24px; font-size:110px; line-height:1; opacity:.08; }
    </style></head><body>
      ${section('01', '#ffffff')}${section('02', '#fdfbf7')}${section('03', '#ffffff')}
    </body></html>`,
    'section-watermarks',
  )
  const documentXml = await xml('word/document.xml')
  const anchors = (documentXml.match(/wp:positionH relativeFrom="column"/g) || []).length
  assert.ok(
    anchors >= 2,
    `later-page watermarks should stay anchored to their sections, got ${anchors}`,
  )
})

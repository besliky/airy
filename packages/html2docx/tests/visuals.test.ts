/**
 * Rasterized fragments: the screenshot pipeline for elements Word cannot
 * express natively — positioned compositions, gradients, rounded cards,
 * nowrap label clusters, icon glyphs, and painted broken-image boxes.
 * One Chrome is shared by every case in this file (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

test('rasterizes bounded relative compositions with positioned text overlays', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .hero { position:relative; width:700px; height:300px; background:#111; }
      .hero-copy { position:absolute; inset:auto 30px 20px; color:white; }
    </style></head><body>
      <header class="hero"><div class="hero-copy"><h1>Overlay title</h1><p>Overlay subtitle</p></div></header>
      <p>Editable body</p>
    </body></html>`,
    'positioned-overlay',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /Overlay title/)
  assert.match(documentXml, /Editable body/)
  assert.ok(
    Object.keys(zip.files).some((name) => name.startsWith('word/media/')),
    'positioned composition should be embedded as screenshot media',
  )
})

test('rasterizes gradient cards nested inside layout rows', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .summary { display:grid; grid-template-columns:1fr 1fr; width:700px; }
      .fee {
        background:linear-gradient(135deg,#ffd700,#ffc700);
        border-radius:12px;
        box-shadow:0 4px 8px rgba(255,170,0,.25);
        padding:14px 20px;
      }
    </style></head><body>
      <section class="summary">
        <div>Event details</div>
        <div class="fee"><strong>¥3,000</strong><div>per person</div></div>
      </section>
    </body></html>`,
    'gradient-row-card',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Event details/)
  assert.match(documentXml, /w:line="\d+" w:lineRule="atLeast"/)
  assert.ok(
    Object.keys(zip.files).some((name) => name.startsWith('word/media/')),
    'gradient fee card should use screenshot media',
  )
})

test('preserves clipped gradient heading text as visual media', async () => {
  const { ir, screenshotText, zip } = await convertHtml(
    `<!doctype html><html><head><style>
      body { background:#0a0f1e; }
      h1 {
        font-size:42px;
        background:linear-gradient(90deg,#e6ecff 30%,#a78bfa 70%);
        -webkit-background-clip:text;
        background-clip:text;
        color:transparent;
      }
    </style></head><body><h1>Gradient report title</h1></body></html>`,
    'gradient-heading-text',
  )
  assert.match(JSON.stringify(ir), /"type":"(?:image|floatimg)"/)
  assert.match(screenshotText, /Gradient report title/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('rasterizes compact rounded cards with exact edge weight', async () => {
  const { zip } = await convertHtml(
    `<!doctype html><html><head><style>
      .card {
        width:600px;
        border:1px solid #e2e8f0;
        border-radius:12px;
        overflow:hidden;
        box-shadow:0 2px 8px -4px rgba(15,23,42,.08);
      }
      .header { background:#059669; color:white; padding:12px 18px; }
      .body { padding:12px 18px; }
    </style></head><body>
      <div class="card"><div class="header">SANGAT KOMPETEN</div><div class="body">Compact assessment card.</div></div>
    </body></html>`,
    'compact-rounded-card',
  )
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('screenshots compact nowrap labels inside table cells', async () => {
  const { zip, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      table { width:600px; border-collapse:collapse; }
      td { width:120px; border:1px solid #ddd; }
      .badge { display:inline-block; white-space:nowrap; color:white; background:#f39c12;
        padding:3px 10px; border-radius:2px; font-size:12px; }
    </style></head><body><table><tr><td>
      <span class="badge">AC・HSPは中</span>／<span class="badge">Fawn特化は空白</span>
    </td></tr></table></body></html>`,
    'nowrap-table-badges',
  )
  assert.match(screenshotText, /AC・HSPは中/)
  assert.match(screenshotText, /Fawn特化は空白/)
  assert.ok(
    Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 1,
    'the compact labels and separator should remain one unbroken cell image',
  )
})

test('rasterizes compact parallel reference blocks', async () => {
  const { ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .references { display:flex; gap:50px; width:500px; }
      .reference { flex:1; }
    </style></head><body><div class="references">
      <div class="reference"><div>Sarah Chen</div><div>GlobalTech / Director</div><div>555-1111</div></div>
      <div class="reference"><div>Michael Thompson</div><div>Pinnacle / VP Operations</div><div>555-2222</div></div>
    </div></body></html>`,
    'parallel-reference-blocks',
  )
  assert.ok(ir.some((node) => node.type === 'image'))
  assert.match(screenshotText, /Michael Thompson/)
})

test('preserves icon-font contacts and compact floated skill bars', async () => {
  const { ir, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      .contact { display:flex; gap:12px; }
      .contact i { display:block; width:20px; height:20px; }
      .contact i::before { content:"☎"; }
      .skills { width:220px; }
      .skill { width:220px; height:50px; }
      .skill .percent { float:right; margin-top:-20px; }
      .track { height:8px; background:#8fb8dd; }
      .progress { width:88%; height:100%; background:#fff; }
      .experience { display:flex; width:420px; gap:12px; }
      .title { flex:1; }
    </style></head><body>
      <div class="contact"><i></i><span>555-123-4567</span></div>
      <div class="skills">
        <div class="skill"><div>Process Optimization <span class="percent">95%</span></div>
          <div class="track"><div class="progress"></div></div></div>
        <div class="skill"><div>Supply Chain Management <span class="percent">88%</span></div>
          <div class="track"><div class="progress"></div></div></div>
      </div>
      <div class="experience"><div>✓</div><div class="title">Senior Coordinator</div><div>2022-2024</div></div>
    </body></html>`,
    'compact-resume-controls',
  )
  assert.match(screenshotText, /Supply Chain Management/)
  assert.ok(ir.filter((node) => node.type === 'image').length >= 1)
  const experience = ir.find(
    (node) =>
      node.type === 'kpirow' &&
      node.cells?.some((cell) => JSON.stringify(cell).includes('Senior Coordinator')),
  )
  assert.ok(experience)
  assert.ok(
    experience.itemWidths[2] >= experience.colWidths[2] + 15,
    'the short date range needs intrinsic-width slack',
  )
})

test('rasterizes painted broken-image placeholders at their authored size', async () => {
  const { zip, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      img { display:block; width:320px; height:240px; background:#e8e5df; border-radius:6px; }
    </style></head><body>
      <img src="file:///definitely-missing-stage-image.png" alt="Stage 1 Front — Earthworks">
    </body></html>`,
    'broken-image-placeholder',
  )
  assert.match(screenshotText, /Stage 1 Front/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('rasterizes compact wrapped flow diagrams without reflowing their steps', async () => {
  const steps = Array.from(
    { length: 7 },
    (_, index) =>
      `${index ? '<div class="arrow">→</div>' : ''}<div class="step">STEP ${index + 1}</div>`,
  ).join('')
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .outer { padding:10px; }
      .flow { display:flex; flex-wrap:wrap; width:420px; gap:8px; }
      .step { width:92px; padding:8px; border:1px solid #999; }
      .arrow { width:18px; }
    </style></head><body><main><section class="outer"><div class="flow">${steps}</div></section></main></body></html>`,
    'wrapped-flow',
  )
  const documentXml = await xml('word/document.xml')
  assert.doesNotMatch(documentXml, /STEP 1[\s\S]*?STEP 7/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('anchors absolute banner decorations inside their card', async () => {
  const { zip, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; padding:32px 80px; }
      .banner { background:#8ba05c; color:#fff; padding:40px 50px; margin:-32px -80px 40px; position:relative; }
      .banner h1 { font-size:3rem; margin:0; text-align:center; }
      .leaf { position:absolute; top:50%; right:50px; transform:translateY(-50%); width:60px; height:60px; }
    </style></head><body>
      <div class="banner">
        <h1>Impact Report</h1>
        <svg class="leaf" viewBox="0 0 24 24" fill="#ffffff"><path d="M4 4h16v16H4z"></path></svg>
      </div>
      <p>Body paragraph after the banner.</p>
    </body></html>`,
    'banner-decoration-anchor',
  )
  const documentXml = await xml('word/document.xml')
  const anchor = documentXml.indexOf('<wp:anchor')
  const table = documentXml.indexOf('<w:tbl>')
  assert.ok(anchor > -1, 'decoration should emit a floating anchor')
  assert.ok(table > -1 && anchor > table, 'anchor should live inside the card table')
  assert.match(documentXml, /wp:positionH relativeFrom="column"/)
  assert.match(documentXml, /wp:positionV relativeFrom="paragraph"/)
  assert.match(documentXml, /behindDoc="0"/)
  assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')))
})

test('screenshots icon-font glyphs that live in pseudo elements', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .contact { display:flex; gap:10px; align-items:center; }
      .icon { width:20px; height:20px; display:block; }
      .icon::before { content:"\\260E"; font-size:16px; }
    </style></head><body>
      <div class="contact"><i class="icon"></i><span>+1 555-789-0123</span></div>
    </body></html>`,
    'icon-font-glyph',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /555-789-0123/)
  const drawing = documentXml.indexOf('<w:drawing>')
  const phone = documentXml.indexOf('555-789-0123')
  assert.ok(drawing > -1 && drawing < phone, 'icon screenshot should precede the phone number')
})

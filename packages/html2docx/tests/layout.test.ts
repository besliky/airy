/**
 * Flowing layout: card borders and padding, heading spacing, section
 * wrappers, KPI/layout rows, and multi-column page layouts — the parts of
 * the document that must stay native and editable.
 * One Chrome is shared by every case in this file (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

test('aligns a padded card border with surrounding paragraphs', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { max-width:880px; margin:0 auto; padding:32px 80px; box-sizing:border-box; }
      .note { border:1px solid #777; padding:14px 16px; background:#fafafa; }
    </style></head><body>
      <p>Before</p>
      <section class="note"><p>Inside</p></section>
      <p>After</p>
    </body></html>`,
    'card-alignment',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:tblInd w:type="dxa" w:w="60"\/>/)
})

test('does not invent padding around zero-padding card headers', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .card { border:3px solid #c0392b; padding:0; width:700px; }
      .header {
        display:flex;
        align-items:center;
        gap:14px;
        padding:12px 18px;
        background:#c0392b;
        color:white;
      }
      .dice { white-space:nowrap; }
    </style></head><body>
      <div class="card"><div class="header"><span class="dice">🎲 目：6</span><span>追加要望カード｜データ取得ミッション</span></div></div>
    </body></html>`,
    'flush-card-header',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(
    documentXml,
    /<w:tcMar><w:top w:type="dxa" w:w="0"\/><w:left w:type="dxa" w:w="0"\/><w:bottom w:type="dxa" w:w="0"\/><w:right w:type="dxa" w:w="0"\/><\/w:tcMar>/,
  )
})

test('keeps bordered inline-block cards shrink-wrapped', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .title { display:inline-block; border:2px solid #000; padding:8px 10px; font-weight:bold; }
    </style></head><body><div class="title">数学的帰納法（不等式証明）</div></body></html><!-- public-hygiene: fixture -->`,
    'shrink-wrapped-card',
  )
  const documentXml = await xml('word/document.xml')
  const width = Number(documentXml.match(/<w:tblW\b[^>]*w:w="(\d+)"[^>]*\/>/)?.[1])
  assert.ok(width > 1000 && width < 6000, `inline card should not span the page, got ${width}`)
})

test('keeps vertical heading borders out of CSS margins', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      h2 {
        border-left:5px solid #148f77;
        padding-left:12px;
        margin-top:40px;
        margin-bottom:12px;
      }
    </style></head><body>
      <p>Before</p>
      <h2>Bordered heading</h2>
      <p>After</p>
    </body></html>`,
    'bordered-heading-spacing',
  )
  const documentXml = await xml('word/document.xml')
  const headingXml = documentXml
    .split('</w:p>')
    .find((paragraph) => paragraph.includes('Bordered heading'))
  assert.ok(headingXml, 'bordered heading paragraph should be present')
  assert.match(headingXml, /w:pBdr/)
  assert.match(headingXml, /w:before="0"/)
  assert.match(headingXml, /w:after="0"/)
})

test('overrides built-in Word spacing for zero-margin headings', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; padding:40px 60px; }
      h1 { margin:0 0 8px; padding:0; }
    </style></head><body><h1>Top aligned report title</h1></body></html>`,
    'zero-heading-margin',
  )
  const documentXml = await xml('word/document.xml')
  const headingXml = documentXml.match(
    /<w:p>[\s\S]*?<w:pStyle w:val="Heading1"\/>[\s\S]*?<\/w:p>/,
  )?.[0]
  assert.ok(headingXml)
  assert.match(headingXml, /w:spacing[^>]*w:before="0"/)
})

test('preserves vertical padding on plain section wrappers', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      section { padding:22px 0 8px; }
      h2 { margin:0; }
    </style></head><body><section><h2>Section heading</h2><p>Section body</p></section></body></html>`,
    'section-wrapper-padding',
  )
  const headingIndex = ir.findIndex(
    (node) => node.type === 'heading' && JSON.stringify(node).includes('Section heading'),
  )
  assert.ok(headingIndex > 0)
  assert.equal(ir[headingIndex - 1].type, 'spacer')
  assert.equal(ir[headingIndex - 1].px, 22)
  assert.equal(ir.at(-1).type, 'spacer')
  assert.equal(ir.at(-1).px, 8)
})

test('keeps margins around thin decorative rules', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .divider { width:700px; height:3px; margin:30px 0; background:#e74c3c; }
    </style></head><body><div class="divider"></div><p>Question</p></body></html>`,
    'thin-rule-spacing',
  )
  const documentXml = await xml('word/document.xml')
  // 30px margins at the margin-floor canvas scale: the default 8px body
  // margin is floored to 24px per side, widening the reported canvas to
  // 826px, so 30 * 15 * (794 / 826) ≈ 433 twips.
  assert.match(documentXml, /<w:spacing w:after="433" w:before="433"/)
})

test('binds a heading and its lede through spacers to the section body', async () => {
  // Regression: heading keepNext only reached the following spacer, so Word
  // orphaned section headings (and heading+subtitle pairs) at page bottoms
  // while the body table/image moved to the next page.
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { font-family: Arial; max-width: 700px; margin: 0 auto; }
      section { margin-top: 48px; }
      h2 { margin: 0 0 8px; }
      .subtitle { color: #667; margin: 0 0 24px; }
      td { border: 1px solid #ccc; padding: 6px 10px; }
    </style></head><body>
      <p>Intro paragraph before any section.</p>
      <section>
        <h2>Yearly Summary</h2>
        <p class="subtitle">Compares uploads and revenue by year</p>
        <table><tr><td>2024</td><td>129</td></tr><tr><td>2025</td><td>287</td></tr></table>
      </section>
    </body></html>`,
    'heading-lede-chain',
  )
  const documentXml = await xml('word/document.xml')
  const paragraphs = documentXml.split('</w:p>')
  const headingIndex = paragraphs.findIndex((p) => p.includes('Yearly Summary'))
  const ledeIndex = paragraphs.findIndex((p) => p.includes('Compares uploads'))
  assert.ok(headingIndex >= 0 && ledeIndex > headingIndex)
  assert.match(paragraphs[headingIndex], /<w:keepNext\/>/)
  assert.match(paragraphs[ledeIndex], /<w:keepNext\/>/)
  // Every paragraph between heading and lede (post-heading spacers) binds on.
  for (let i = headingIndex + 1; i < ledeIndex; i++) {
    assert.match(paragraphs[i], /<w:keepNext\/>/)
  }
  // The intro paragraph before the section must NOT join any keep chain.
  const introIndex = paragraphs.findIndex((p) => p.includes('Intro paragraph'))
  assert.doesNotMatch(paragraphs[introIndex], /<w:keepNext\/>/)
})

test('subtracts page-column padding from nested layout width', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .resume { display:grid; grid-template-columns:278px 516px; width:794px; }
      aside { min-height:800px; background:#1e3a5f; padding:30px; }
      main { min-height:800px; padding:40px 50px; }
      .skill { display:flex; justify-content:space-between; }
      .track { width:187px; height:8px; background:#ff6633; }
    </style></head><body>
      <div class="resume">
        <aside>Sidebar</aside>
        <main><div class="skill"><span>Planning</span><div class="track"></div></div></main>
      </div>
    </body></html>`,
    'padded-page-columns',
  )
  const documentXml = await xml('word/document.xml')
  const tableWidths = [...documentXml.matchAll(/<w:tblW w:type="dxa" w:w="(\d+)"\/>/g)].map(
    (match) => Number(match[1]),
  )
  // main column minus its 50px paddings at the margin-floor canvas scale:
  // (516 - 100) * 15 * (794 / 842) ≈ 5885 dxa; unpadded would be ~7300.
  assert.ok(
    tableWidths.some((width) => width >= 5700 && width <= 6100),
    `nested main-column table should use padded width, got ${tableWidths.join(', ')}`,
  )
})

test('keeps a full-width resume header above tall grid columns', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .resume {
        display:grid;
        grid-template-columns:3fr 2fr;
        gap:24px;
        width:760px;
        position:relative;
      }
      header { grid-column:1 / -1; }
      .left, .right { min-height:850px; }
      .decoration { position:absolute; right:0; bottom:0; width:80px; height:80px; background:#eef; }
    </style></head><body>
      <div class="resume">
        <div class="decoration"></div>
        <header><h1>Sarah Mitchell</h1><p>Senior Brand Designer</p></header>
        <main class="left"><h2>Work Experience</h2><p>Left column</p></main>
        <aside class="right"><h2>About</h2><p>Right column</p></aside>
      </div>
    </body></html>`,
    'headed-page-columns',
  )
  const pageColumns = ir.find((node) => node.type === 'table' && node.pageColumns)
  assert.ok(pageColumns, 'the two tall columns should remain a page-column table')
  assert.deepEqual(pageColumns.colWidths.length, 2)
  assert.ok(
    ir.some(
      (node) => node.type === 'heading' && node.runs?.some((run) => run.text.includes('Sarah')),
    ),
    'the spanning header should stay before the column table',
  )
})

test('paints the cell edge behind a top-aligned visual page-column header', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      * { box-sizing:border-box; }
      body { margin:0; background:#f4f7fb; }
      .layout { display:grid; grid-template-columns:220px 574px; width:794px; }
      aside { min-height:800px; background:#0d2137; }
      main { min-height:800px; padding:0 48px; }
      .hero { height:140px; margin:0 -48px; background:linear-gradient(135deg,#0d2137,#1e4d80); color:white; }
    </style></head><body>
      <div class="layout">
        <aside>Navigation</aside>
        <main><div class="hero"><h1>SDK Reference</h1></div><p>Body text</p></main>
      </div>
    </body></html>`,
    'page-column-visual-header',
  )
  const pageColumns = ir.find((node) => node.type === 'table' && node.pageColumns)
  assert.ok(pageColumns)
  assert.equal(pageColumns.rows[0].cells[0].gapAfterPx, 0)
  assert.equal(pageColumns.rows[0].cells[1].children[0].type, 'image')
  assert.equal(pageColumns.rows[0].cells[1].children[0].align, 'left')
  assert.equal(pageColumns.rows[0].cells[1].children[0].bleedLeftPx, 48)
  assert.equal(pageColumns.rows[0].cells[1].children[0].topEdgeFill, '0D2137')
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /<w:top w:val="single" w:color="0D2137" w:sz="48" w:space="0"\/>/)
})

test('preserves tall multi-row table grids as side-by-side visual rows', async () => {
  const { ir, zip, screenshotText } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .ranking-grid {
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:16px;
        width:1200px;
      }
      .table-wrap { height:760px; border:1px solid #ddd; overflow:hidden; }
      table { width:100%; border-collapse:collapse; }
      td { height:60px; border-bottom:1px solid #ddd; }
    </style></head><body>
      <div class="ranking-grid">
        <div class="table-wrap"><table><tr><td>Views table</td></tr></table></div>
        <div class="table-wrap"><table><tr><td>Watch time table</td></tr></table></div>
        <div class="table-wrap"><table><tr><td>Subscribers table</td></tr></table></div>
        <div class="table-wrap"><table><tr><td>Revenue table</td></tr></table></div>
      </div>
    </body></html>`,
    'large-table-grid',
  )
  const grid = ir.find((node) => node.type === 'table' && node.largeTableGrid)
  assert.ok(grid, 'the tall table grid should retain an explicit layout wrapper')
  const contentRows = grid.rows.filter((row) => !row.gapRow)
  assert.equal(contentRows.length, 2)
  assert.ok(
    contentRows.every(
      (row) =>
        row.cells.filter((cell) => cell.children?.length).length === 2 &&
        row.cells
          .filter((cell) => cell.children?.length)
          .every((cell) => cell.children[0].type === 'image'),
    ),
  )
  assert.equal(
    ir.some((node) => node.pageColumns),
    false,
  )
  assert.match(screenshotText, /Views table/)
  assert.match(screenshotText, /Revenue table/)
  assert.ok(Object.keys(zip.files).filter((name) => name.startsWith('word/media/')).length >= 4)
})

test('rasterizes a page column containing positioned decorations', async () => {
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .resume { display:flex; gap:24px; width:760px; }
      .left, .right { min-height:850px; }
      .left { width:260px; }
      .right { position:relative; width:476px; padding-top:180px; }
      .photo { position:absolute; top:0; right:40px; width:140px; height:140px; background:#b88; }
      .dot { position:absolute; left:4px; top:300px; width:10px; height:10px; background:#843; }
    </style></head><body>
      <div class="resume">
        <aside class="left"><h1>Naval Kishor</h1><p>Education and skills</p></aside>
        <main class="right">
          <div class="photo"></div><div class="dot"></div>
          <h2>Professional Summary</h2><p>Summary text</p>
          <h2>Work Experience</h2><p>Experience text</p>
        </main>
      </div>
    </body></html>`,
    'positioned-page-column',
  )
  const pageColumns = ir.find((node) => node.type === 'table' && node.pageColumns)
  assert.ok(pageColumns)
  assert.equal(pageColumns.rows[0].cells[1].children.length, 1)
  assert.equal(pageColumns.rows[0].cells[1].children[0].type, 'image')
  assert.equal(JSON.stringify(pageColumns).includes('floatimg'), false)
})

test('keeps a row of 3+ tall equal pricing cards side by side', async () => {
  const card = (name, price) =>
    `<div class="pkg"><h3>${name}</h3><p>R$ ${price}</p>
     <ul><li>Feature one</li><li>Feature two</li><li>Feature three</li></ul></div>`
  const { ir } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; }
      .grid { display:grid; grid-template-columns:repeat(5, 1fr); gap:0; }
      .pkg { border-right:1px solid #ddd; padding:40px 20px; height:560px; box-sizing:border-box; }
    </style></head><body>
      <div class="grid">
        ${card('Kit Essencial', 490)}${card('Kit Comunicacao', 890)}${card('Kit Lancamento', 1490)}
        ${card('Social Business', 2490)}${card('Social Premium', 3990)}
      </div>
    </body></html>`,
    'parallel-pricing-cards',
  )
  const row = ir.find((node) => node.type === 'kpirow' && (node.cells || []).length === 5)
  assert.ok(row, 'five tall equal columns should classify as one kpirow')
})

test('keeps a KPI row horizontal inside a card nested in a layout row', async () => {
  const { ir, xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .hero { display:grid; grid-template-columns:1.4fr .9fr; gap:18px; width:900px; }
      .panel { background:#fff; border:1px solid #ddd; border-radius:12px; padding:22px; }
      .stats { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; }
      .stat { background:#f8fafc; border:1px solid #ddd; padding:16px; }
      .value { font-size:28px; font-weight:800; }
    </style></head><body>
      <div class="hero">
        <div class="panel">
          <h2>Highlights</h2>
          <div class="stats">
            <div class="stat"><div>Views</div><div class="value">48,433,718</div></div>
            <div class="stat"><div>Hours</div><div class="value">972,931.9915</div></div>
            <div class="stat"><div>Revenue</div><div class="value">3,988,282.297</div></div>
          </div>
        </div>
        <div class="panel"><h2>Summary</h2><p>Supporting details</p></div>
      </div>
    </body></html>`,
    'nested-kpi-row',
  )
  const outerRow = ir.find(
    (node) => node.type === 'kpirow' && JSON.stringify(node).includes('Highlights'),
  )
  assert.ok(outerRow)
  assert.ok(JSON.stringify(outerRow).includes('"cells":[{"children"'))
  const documentXml = await xml('word/document.xml')
  assert.ok(
    (documentXml.match(/<w:tbl>/g) || []).length >= 1,
    'outer layout row should remain a table',
  )
  // the lone panel card hoists its border onto the outer row's cell so
  // sibling panels share the row height (flex-stretch equal heights)
  assert.ok(
    /<w:tcBorders>(?:(?!<\/w:tcBorders>)[^])*?w:color="DDDDDD"/.test(documentXml),
    'panel border should sit on the outer row cell',
  )
})

test('carries flattened section shading through spacers and nested blocks', async () => {
  const paras = Array.from({ length: 10 }, (_, i) => `<p>Filler paragraph ${i + 1}.</p>`).join('')
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { margin:0; background:#efe9dd; }
      section { background:#ffffff; padding:40px; }
      .cols { display:flex; gap:24px; }
    </style></head><body>
      <section>
        <h2>White section</h2>
        ${paras}
        <div class="cols"><div>Column A</div><div>Column B</div></div>
        ${paras}
      </section>
    </body></html>`,
    'flatten-shading',
  )
  const documentXml = await xml('word/document.xml')
  const fills = (documentXml.match(/w:fill="FFFFFF"/g) || []).length
  assert.ok(fills >= 10, `flattened white section should shade its blocks, got ${fills}`)
  // the zero-height spacer paragraphs must be shaded too, not just text paras
  assert.match(documentXml, /<w:shd[^>]*w:fill="FFFFFF"[^>]*\/><w:spacing[^>]*w:lineRule="exact"/)
})

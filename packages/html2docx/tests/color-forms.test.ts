/**
 * Color compositing and form scaffolding: rgba tints over backgrounds,
 * white cards on gradient pages, and empty answer boxes / fill-in lines /
 * compact footer label rows that must survive as native Word constructs.
 * One Chrome is shared by every case in this file (see support/convert.ts).
 */
import assert from 'node:assert/strict'
import { test } from 'vitest'

import { setupFileConversion } from './support/convert'

const convertHtml = setupFileConversion()

test('keeps white cards over gradient page backgrounds', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 24px;
      }
      .content-box {
        background: white;
        border-radius: 8px;
        padding: 20px;
        margin: 20px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,.1);
        border: 1px solid rgba(0,0,0,.08);
      }
    </style></head><body>
      <div class="content-box">On this day the first party agrees to provide comprehensive
      maintenance services covering documentation, training, monitoring visits,
      and on-site inspection within forty eight hours of remote failure.</div>
      <div class="content-box">The second party shall maintain strict confidentiality regarding all
      proprietary business information and financial data of the first party under
      this maintenance cooperation agreement letter for two full years.</div>
    </body></html>`,
    'white-card-on-gradient',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /first party agrees/)
  assert.match(documentXml, /second party shall maintain/)
  const whiteFills = (documentXml.match(/w:fill="FFFFFF"/g) || []).length
  assert.ok(whiteFills >= 2, `expected white card shading, got ${whiteFills}`)
})

test('composites semi-transparent backgrounds onto white', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .note {
        background:rgba(196,30,58,.08);
        color:#c41e3a;
        padding:8px 12px;
      }
    </style></head><body>
      <p class="note">Tinted note</p>
    </body></html>`,
    'rgba-tint',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Tinted note/)
  // Solid accent red must not be used as the fill; the tint composites to ~FAEDEF.
  assert.doesNotMatch(documentXml, /w:fill="C41E3A"/)
  assert.match(documentXml, /w:fill="FAEDEF"/)
})

test('composites semi-transparent table stripes onto a dark panel', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { background:#0a0f1e; color:#aab4d4; }
      table { width:600px; background:#10172e; border-collapse:collapse; }
      th { background:rgba(16,23,46,.85); color:#cfd8f5; padding:8px; }
      td { color:#aab4d4; padding:8px; }
      tbody tr:nth-child(odd) { background:rgba(15,21,40,.4); }
    </style></head><body>
      <table>
        <thead><tr><th>Tier</th><th>Audience</th></tr></thead>
        <tbody>
          <tr><td>Entry</td><td>Students</td></tr>
          <tr><td>Premium</td><td>Creators</td></tr>
        </tbody>
      </table>
    </body></html>`,
    'dark-table-stripes',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Entry/)
  assert.doesNotMatch(documentXml, /w:fill="9FA1A9"/)
  assert.match(documentXml, /w:fill="10162C"|w:fill="10172E"/)
})

test('keeps empty bordered answer boxes with authored height', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .box { border:1px solid #d0d5dd; background:#fcfcfd; padding:12px; }
      .field-label { display:block; font-weight:700; font-size:12px; }
      .answer-box {
        min-height:58px;
        border:1px solid #d0d5dd;
        background:#fff;
        margin-top:8px;
      }
    </style></head><body>
      <div class="box">
        <span class="field-label">WHAT SHOULD THE AI HELP CUSTOMERS DO?</span>
        <div class="answer-box"></div>
      </div>
    </body></html>`,
    'answer-box',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /WHAT SHOULD THE AI HELP/)
  // Answer box becomes a fixed-height form cell, not a collapsed spacer.
  assert.match(documentXml, /w:sdt/)
  assert.ok(
    /w:trHeight[^>]*w:val="[1-9][0-9]{2,}"/.test(documentXml),
    'answer box row should keep a substantial height',
  )
})

test('keeps empty underline fill-in fields', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .field-line { border-bottom:1px solid #000; height:20px; width:280px; }
      .dotted-line {
        border-bottom:1px dotted black;
        display:inline-block;
        min-width:60px;
        margin:0 3px;
      }
    </style></head><body>
      <p>Full Name:</p>
      <div class="field-line"></div>
      <p>On this day <span class="dotted-line"></span> date</p>
    </body></html>`,
    'fill-in-lines',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /Full Name/)
  assert.match(documentXml, /w:bottom[^>]*w:color="000000"/)
  assert.match(documentXml, /w:u w:val="dotted"/)
  assert.match(documentXml, /On this day/)
})

test('keeps short flex meta labels on one line', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      body { max-width:880px; margin:0 auto; padding:32px 80px; }
      .meta-row { display:flex; gap:28px; font-size:13px; }
      .meta-row span strong { font-weight:600; }
    </style></head><body>
      <h1>Notice title</h1>
      <div class="meta-row">
        <span><strong>수신</strong> SM 전원</span>
        <span><strong>발신</strong> 박규하 RM</span>
        <span><strong>일자</strong> 2025년 7월 11일</span>
      </div>
    </body></html>`,
    'flex-meta-row',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /수신/)
  assert.match(documentXml, /SM 전원/)
  assert.match(documentXml, /발신/)
  assert.match(documentXml, /박규하 RM/)
  // Each meta cell should stay a single paragraph (no mid-label wrap split
  // into a second run/paragraph inside the cell).
  assert.doesNotMatch(documentXml, /수신[^<]{0,20}SM<\/w:t><\/w:r><\/w:p><w:p>/)
})

test('keeps long email contact cells on one line in flex footers', async () => {
  const { xml } = await convertHtml(
    `<!doctype html><html><head><style>
      .footer-contact {
        display:flex;
        justify-content:space-between;
        align-items:center;
        border-top:2px solid #d4af37;
        padding-top:12px;
        font-size:15px;
        color:#8b6f47;
      }
    </style></head><body>
      <p>Resume body</p>
      <div class="footer-contact">
        <span>info@candidatename.email.com</span>
        <span>(555) 456-7890</span>
        <span>1234 Professional Drive, Business City, ST</span>
      </div>
    </body></html>`,
    'email-footer',
  )
  const documentXml = await xml('word/document.xml')
  assert.match(documentXml, /info@candidatename\.email\.com/)
  // Email must live in a reasonably wide cell (not ~227px that wraps "m").
  const widths = [...documentXml.matchAll(/<w:tcW[^>]*\bw:w="(\d+)"/g)].map((m) => Number(m[1]))
  assert.ok(widths.length >= 3, `expected footer table cells, got ${widths.length}`)
  // ~280px email cell ≈ 4200 twips; gaps may shrink but content stays wide.
  assert.ok(
    widths.some((w) => w >= 3800),
    `expected a wide contact cell for the email, got ${widths.join(',')}`,
  )
})

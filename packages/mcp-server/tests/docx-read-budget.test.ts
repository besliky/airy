// BUG-1684 regression: read_document answers stay inside the 30k-character
// budget on BOTH paths. The default read used to answer a 42k-block document
// with ~800k characters (the single-pass elide kept 2/3 of the overview
// lines), and the blocks/range path appended the FULL overview of all blocks,
// so even a five-block range answered with ~2.9MB. The response contract
// (header, overview lines, stats, selected HTML) must hold for valid small
// files exactly as before — only the sizes are bounded.
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { DocxSession } from '../src/docx/session.js'
import { buildFixtureDocx, buildManyBlocksDocx } from './helpers/docx-fixture.js'

/** the documented honest answer budget (CONTEXT_MAX_CHARS in the session) */
const BUDGET_BYTES = 30_000

let root: string
let smallPath: string
let bigPath: string
let small: DocxSession
let big: DocxSession

beforeEach(async () => {
  root = await mkdtemp(join(tmpdir(), 'airy-docx-budget-'))
  smallPath = join(root, 'report.docx')
  await writeFile(smallPath, await buildFixtureDocx())
  bigPath = join(root, 'flood.docx')
  // mirrors the audited 42k-block corpus whose overview alone is ~800k chars
  await writeFile(bigPath, await buildManyBlocksDocx(42_000))
  small = await DocxSession.open(smallPath, root)
  big = await DocxSession.open(bigPath, root)
})

afterEach(async () => {
  await rm(root, { recursive: true, force: true })
})

describe('docx read_document budget (BUG-1684)', () => {
  it('default read of a 42k-block document stays inside the budget', () => {
    const out = big.readDocument()
    expect(Buffer.byteLength(out, 'utf8')).toBeLessThanOrEqual(BUDGET_BYTES)
    // the elide ladder keeps the answer honest: both ends and the marker survive
    expect(out).toContain('The document has 42000 blocks')
    // previews are tightened, so the line ends in the clip ellipsis; the
    // block indexes stay verifiable at both ends
    expect(out).toMatch(/^0\|p\|Block 0 /m)
    expect(out).toMatch(/^41999\|p\|Block 41999 /m)
    expect(out).toMatch(/blocks elided here; numbering is continuous/)
    expect(out).toContain('Full-text stats:')
  })

  it('range read of five far-apart blocks stays inside the budget and keeps the payload', () => {
    const out = big.readDocument({ range: { start: 41_000, end: 41_004 } })
    expect(Buffer.byteLength(out, 'utf8')).toBeLessThanOrEqual(BUDGET_BYTES)
    // the selected blocks ride along in full restricted HTML
    expect(out).toContain('<p>Block 41000 lorem ipsum dolor</p>')
    expect(out).toContain('<p>Block 41004 lorem ipsum dolor</p>')
    expect(out).toContain('Selected block content (restricted HTML):')
    // the formerly-unbounded companion overview is now elided, not dropped
    expect(out).toContain('The document has 42000 blocks')
    expect(out).toMatch(/blocks elided here; numbering is continuous/)
  })

  it('max-span range read (10000 blocks of HTML) is clipped to the budget with a note', () => {
    const out = big.readDocument({ range: { start: 0, end: 9_999 } })
    expect(Buffer.byteLength(out, 'utf8')).toBeLessThanOrEqual(BUDGET_BYTES)
    expect(out).toContain('output truncated')
    expect(out).toContain('request a narrower block range')
    // the payload still leads with the first selected blocks
    expect(out).toContain('<p>Block 0 lorem ipsum dolor</p>')
  })

  it('block selection on a huge document keeps the payload and caps the overview', () => {
    const out = big.readDocument({ blocks: [0, 41_999] })
    expect(Buffer.byteLength(out, 'utf8')).toBeLessThanOrEqual(BUDGET_BYTES)
    expect(out).toContain('<p>Block 0 lorem ipsum dolor</p>')
    expect(out).toContain('<p>Block 41999 lorem ipsum dolor</p>')
  })

  it('small document default read keeps the full overview contract', () => {
    const out = small.readDocument()
    expect(out).toContain('The document has 7 blocks')
    expect(out).toMatch(/^0\|h1\|Quarterly Report$/m)
    expect(out).toMatch(/^1\|p\|Revenue grew by 12 percent year over year\.$/m)
    expect(out).toMatch(/^2\|li\|First bullet item$/m)
    expect(out).toMatch(/^5\|table\|Region \| Sales \/ East \| 4200$/m)
    expect(out).toContain('Full-text stats:')
    // nothing elided on a document that fits
    expect(out).not.toContain('elided here')
  })

  it('small document selection read keeps overview AND full HTML (contract unchanged)', () => {
    const out = small.readDocument({ blocks: [0, 1] })
    expect(out).toContain('<h1>Quarterly Report</h1>')
    expect(out).toContain('<p>Revenue grew by <strong>12 percent</strong> year over year.</p>')
    // overview lines and stats still precede the selection for small files
    expect(out).toMatch(/^0\|h1\|Quarterly Report$/m)
    expect(out).toMatch(/^2\|li\|First bullet item$/m)
    expect(out).toContain('Full-text stats:')
    expect(out).not.toContain('elided here')
    expect(out).not.toContain('output truncated at')
  })
})

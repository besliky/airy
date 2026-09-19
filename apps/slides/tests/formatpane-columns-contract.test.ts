import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

/**
 * TEST-1104: the FormatPane text-column controls (PAR-301) must keep writing
 * through the document ops — the count/gap fields route numCol/spcCol into
 * onTextBodyProps, which App.tsx forwards as a setTextBodyProps
 * document op (the op-side validation — numCol 1-13, spcCol in EMU — is
 * already pinned in edit-ops.test.ts). The unpinned link was the UI wiring:
 * a pane refactor that mutates shape.text directly, drops the 0.5" default
 * gap, or bypasses the op path would serialize nothing to the document and
 * no test would notice until a save-reopen lost the columns.
 *
 * Mechanical traits, pinned on the two files of the chain:
 *   - FormatPane.tsx: the count picker writes { numCol: n, ... } through
 *     onTextBodyProps and carries PowerPoint's 0.5" (457200 EMU) default
 *     gap when going multi-column with no gap set; the gap commit and nudge
 *     write spcCol in EMU (cm * 360000); the gap row renders only for
 *     numCol > 1; and nothing assigns shape.text.(numCol|spcCol) directly.
 *   - App.tsx: the onTextBodyProps prop forwards (id, props) verbatim into
 *     window.slidesApi.setTextBodyProps({ slideIndex, sourceId, props }).
 */
const PANE = join(__dirname, '../src/renderer/components/FormatPane.tsx')
const APP = join(__dirname, '../src/renderer/App.tsx')

describe('FormatPane column fields write numCol/spcCol through ops (TEST-1104)', () => {
  const pane = readFileSync(PANE, 'utf8')

  it('the count picker routes through onTextBodyProps with the 0.5" default gap', () => {
    expect(pane, 'count write must go through onTextBodyProps').toMatch(
      /onTextBodyProps\(\s*node\.sourceId,\s*\{\s*numCol:\s*n,/s,
    )
    expect(pane, 'multi-column presets must carry the 457200 EMU default gap').toMatch(
      /n\s*>\s*1\s*&&\s*!\(shape\.text!\.spcCol\s*\?\?\s*0\)\s*\?\s*\{\s*spcCol:\s*457200\s*\}/,
    )
  })

  it('the gap field commits and nudges spcCol in EMU, only when multi-column', () => {
    const gapWrites = pane.match(
      /onTextBodyProps\(\s*node\.sourceId,\s*\{\s*spcCol:\s*Math\.round\((?:v|next)\s*\*\s*360000\)\s*\}\)/gs,
    )
    // exactly two: the typed commit and the stepper nudge
    expect(gapWrites?.length).toBe(2)
    expect(pane, 'the gap row renders only for numCol > 1').toMatch(
      /\{\(shape\.text\.numCol\s*\?\?\s*1\)\s*>\s*1\s*&&/,
    )
  })

  it('never mutates the column fields on the shape directly', () => {
    expect(pane.match(/shape\.text!?\.(?:numCol|spcCol)\s*=[^=]/)).toBeNull()
  })

  it('App forwards onTextBodyProps into the setTextBodyProps document op', () => {
    const app = readFileSync(APP, 'utf8')
    expect(app).toMatch(
      /onTextBodyProps=\{\(id,\s*props\)\s*=>\s*void window\.slidesApi\s*\.\s*setTextBodyProps\(\{\s*slideIndex:\s*current,\s*sourceId:\s*id,\s*props\s*\}\)/s,
    )
  })
})

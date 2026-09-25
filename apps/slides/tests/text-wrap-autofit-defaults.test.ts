/**
 * OBS-1728 regression: text-box wrap default and cache-less normAutofit render.
 *
 * The 2026-09-25 slides-editing audit (SL-EDIT-5) reported that text boxes "without an
 * explicit wrap" render as a single overflowing line and that a bare <a:normAutofit/>
 * renders at 100% with overflow. Both asks already behave correctly on main; the audit's
 * corpus artifact carries explicit wrap="none" on every box (python-pptx's add_textbox
 * template hard-codes it), for which single-line overflow is the PowerPoint-correct
 * render. These tests lock the actual OOXML-default behaviors so they cannot regress:
 *   - a bodyPr without a wrap attribute parses as square wrap and lays out at box width;
 *   - wrap="none" parses as no wrap and stays single-line (the audited corpus bytes);
 *   - a bare normAutofit (no stored fontScale) steps down the shrink ladder on plain
 *     render when the content overflows the box height (stored scales stay authoritative).
 */
import { describe, it, expect } from 'vitest'
import {
  openPptx,
  savePptx,
  addElement,
  createBlankPptx,
  type TextBody,
} from '@airy-office/pptx-engine'
import { layoutText, makeViewport, HeuristicMetrics } from '@airy-office/pptx-render'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280) // scale = 1
const metrics = new HeuristicMetrics()

const LONG =
  'AUTOSIZE-C bottom cut large text that must shrink a lot to fit somewhere inside the box'

const BOX_W = 3429000 / 12700 // 3.75in box in px
const BOX_H = 1143000 / 12700
const INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 }

const layoutBody = (body: TextBody) =>
  layoutText({ body, boxWidthPx: BOX_W, boxHeightPx: BOX_H, metrics, vp })

/** Build a deck with one text box, rewrite its slide XML, and re-open it (parse-level probe). */
async function parsedBodyPr(bodyPrXml: string): Promise<TextBody> {
  const opened = await openPptx(await createBlankPptx())
  addElement(opened.deck.slides[0]!, {
    kind: 'textbox',
    offset: { x: 100000, y: 100000, cx: 3429000, cy: 1143000 },
    paragraphs: [{ runs: [{ text: LONG, fontSize: 40 }] }],
  })
  const bytes = await savePptx(opened)
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(bytes)
  const path = opened.deck.slides[0]!.path
  let xml = await zip.file(path)!.async('string')
  // addElement writes <a:bodyPr wrap="square" rtlCol="0"/> — swap in the shape under test
  xml = xml.replace('<a:bodyPr wrap="square" rtlCol="0"/>', bodyPrXml)
  zip.file(path, xml)
  const reopened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
  const el = reopened.deck.slides[0]!.elements.find((e) => e.type === 'text')
  if (!el || el.type !== 'text' || !el.text) throw new Error('text box missing after reparse')
  return el.text
}

describe('OBS-1728: OOXML-default wrap renders wrapped at box width', () => {
  it('parse: a bodyPr without a wrap attribute means square wrap (not none)', async () => {
    const body = await parsedBodyPr('<a:bodyPr rtlCol="0"/>')
    // The model default is wrap: square — a missing attribute must not read as "no wrap"
    expect(body.wrap).toBe(true)
  })

  it('layout: default wrap breaks the text into lines that fit the box width', () => {
    const out = layoutBody({
      paragraphs: [{ runs: [{ text: LONG, fontSize: 40 }] }],
      insets: INSETS,
    })
    expect(out.lines.length).toBeGreaterThan(1)
    for (const line of out.lines) {
      expect(line.runs.reduce((w, r) => w + r.widthPx, 0)).toBeLessThanOrEqual(BOX_W)
    }
  })
})

describe('OBS-1728: explicit wrap="none" (the audited corpus bytes) stays single-line', () => {
  it('parse + layout: wrap="none" is honored as no wrapping with overflow', async () => {
    const body = await parsedBodyPr('<a:bodyPr wrap="none" rtlCol="0"/>')
    expect(body.wrap).toBe(false)
    const out = layoutBody({
      paragraphs: [{ runs: [{ text: LONG, fontSize: 40 }] }],
      insets: INSETS,
      wrap: false,
    })
    expect(out.lines).toHaveLength(1)
  })
})

describe('OBS-1728: bare normAutofit without a stored fontScale', () => {
  it('parse: <a:normAutofit/> maps to shrink autofit with no stored scale', async () => {
    const body = await parsedBodyPr('<a:bodyPr rtlCol="0"><a:normAutofit/></a:bodyPr>')
    expect(body.autofit).toBe('shrink')
    expect(body.fontScale).toBeUndefined()
  })

  it('plain render steps down the shrink ladder when the content overflows the box height', () => {
    // Cache-less box (no stored fontScale): the ladder serves the plain render too
    const out = layoutBody({
      paragraphs: [{ runs: [{ text: LONG, fontSize: 40 }] }],
      insets: INSETS,
      autofit: 'shrink',
    })
    expect(out.fontScale).toBeLessThan(1)
    // A stored scale stays authoritative on plain render (PowerPoint-on-open semantics)
    const stored = layoutBody({
      paragraphs: [{ runs: [{ text: LONG, fontSize: 40 }] }],
      insets: INSETS,
      autofit: 'shrink',
      fontScale: 0.625,
    })
    expect(stored.fontScale).toBeCloseTo(0.625, 3)
  })

  it('a single line that fits the box height is not shrunk (the audit box C shape: wrap=none + bare normAutofit)', () => {
    // With wrap="none" the 40pt line fits the box HEIGHT, so there is nothing to shrink —
    // same verdict PowerPoint gives the audited autosize-edge.pptx bytes
    const out = layoutBody({
      paragraphs: [{ runs: [{ text: LONG, fontSize: 40 }] }],
      insets: INSETS,
      autofit: 'shrink',
      wrap: false,
    })
    expect(out.lines).toHaveLength(1)
    expect(out.fontScale).toBe(1)
  })
})

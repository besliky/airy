/**
 * BUG-1402 regression: the product (DOM-measurement) footnote reservation must
 * charge each entry paragraph's spacing, not just its line boxes.
 *
 * The parity harness (pagination-parity.test.ts) reserves through
 * estimateFootnoteHeight, whose computeLineMetrics model carries the note
 * style's space before/after — that engine model matches Word. The product
 * path (App.tsx noteRenderInfoOf) instead measures the entry row in the DOM
 * (line boxes only, the renderers' CSS line-height) and used to reserve the
 * bare row height. On doc06 (20 single-line notes on a Normal chain,
 * space-after 160tw) that under-reserved 10.67px per note: page 1 held ~139px
 * more body text than Word and pulled extra paragraphs past Word's break.
 *
 * This test rebuilds the product reservation (wrapped rows × noteLineHeightPx,
 * topped up by noteReservedHeightPx) and asserts the corpus engine paginates
 * doc06 exactly like the recorded Word baseline.
 */
import { describe, expect, it, beforeAll } from 'vitest'
import { readFileSync, existsSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { parseDocx, readSections } from '@airy-office/docx-engine'
import {
  computeSectionedSlicesF2,
  sectionGeoms,
  sectionColGeom,
  type BlockBox,
} from '../src/renderer/pagination'
import {
  computeLineMetrics,
  noteLineHeightPx,
  noteReservedHeightPx,
  noteRunStyle,
  resolveNoteStyle,
  type NoteStyleOpts,
} from '../src/renderer/line-metrics'
import { loBaselineMetrics } from './helpers/lo-fonts'

const __dirname = dirname(fileURLToPath(import.meta.url))
const CORPUS_DIR = join(__dirname, 'pagination-corpus/docx')
const WORD_BASELINE_FILE = join(__dirname, 'pagination-corpus/baseline-word.json')

const metrics = loBaselineMetrics()

interface WordBaselineEntry {
  pages: number
  pageStarts: string[]
  pageEnds: string[]
}

let baseline: WordBaselineEntry | undefined

describe('DOM-measured footnote reservation paginates doc06 like Word (BUG-1402)', () => {
  beforeAll(() => {
    if (existsSync(WORD_BASELINE_FILE)) {
      const raw = JSON.parse(readFileSync(WORD_BASELINE_FILE, 'utf-8'))
      const entry = raw['06-with-footnotes']
      if (entry && !entry.error && entry.pages > 0) baseline = entry
    }
  })

  it.skipIf(!existsSync(join(CORPUS_DIR, '06-with-footnotes.docx')))(
    'noteReservedHeightPx charges the entry spacing on top of the measured row',
    async () => {
      const bytes = readFileSync(join(CORPUS_DIR, '06-with-footnotes.docx'))
      const parsed = await parseDocx(new Uint8Array(bytes))
      const sections = readSections(parsed)
      const geoms = sectionGeoms(sections)
      const contentWidthPx = sectionColGeom(sections[0]).colWidthPx
      const docGrid = sections[0].settings.docGrid

      // note style resolved like App.tsx noteStyleOf: style chain + run style
      const styleOf = (fn: (typeof parsed.footnotes)[number]): NoteStyleOpts => ({
        ...resolveNoteStyle(parsed, fn.styleId, fn.spacing),
        ...noteRunStyle(fn.richParas),
      })
      const footnoteById = new Map(parsed.footnotes.map((f) => [f.id, f]))

      // product-model reservation: wrapped line count × CSS line-height row,
      // topped up with the entry paragraph's spacing (BUG-1402 fix)
      const reservedHeightOf = (fn: (typeof parsed.footnotes)[number]): number => {
        const style = styleOf(fn)
        const lineH = noteLineHeightPx(docGrid, style)
        const lines = computeLineMetrics({
          runs: [{ text: fn.text }],
          availWidthPx: contentWidthPx - 40,
          docGrid,
          defaultFontSizePt: style.sizeHalfPoints ? style.sizeHalfPoints / 2 : 10,
          ...(style.fontFamily ? { defaultFontFamily: style.fontFamily } : {}),
          ...(style.lineRule ? { lineRule: style.lineRule } : {}),
          ...(style.lineRawTwips ? { lineRawTwips: style.lineRawTwips } : {}),
          metrics,
        }).lineBoxes.length
        return noteReservedHeightPx(lines * lineH, style)
      }
      // bare-row model (the BUG-1402 defect): line boxes only, no spacing
      const bareHeightOf = (fn: (typeof parsed.footnotes)[number]): number => {
        const style = styleOf(fn)
        const lineH = noteLineHeightPx(docGrid, style)
        const lines = computeLineMetrics({
          runs: [{ text: fn.text }],
          availWidthPx: contentWidthPx - 40,
          docGrid,
          defaultFontSizePt: style.sizeHalfPoints ? style.sizeHalfPoints / 2 : 10,
          ...(style.fontFamily ? { defaultFontFamily: style.fontFamily } : {}),
          ...(style.lineRule ? { lineRule: style.lineRule } : {}),
          ...(style.lineRawTwips ? { lineRawTwips: style.lineRawTwips } : {}),
          metrics,
        }).lineBoxes.length
        return lines * lineH
      }

      // the spacing term is present and material (Normal chain: after=160tw)
      const anyNote = parsed.footnotes.find((f) => Number(f.id) > 0)!
      const style = styleOf(anyNote)
      expect(style.spaceAfterTwips).toBe(160)
      expect(reservedHeightOf(anyNote) - bareHeightOf(anyNote)).toBeCloseTo((160 * 96) / 1440, 4)

      // paginate with a given per-note height model (product blocks like the
      // parity harness: computeLineMetrics bodies + per-reference note bands)
      const paginateWith = (
        heightOf: (fn: (typeof parsed.footnotes)[number]) => number,
      ): string[] => {
        const blocks: BlockBox[] = []
        const blockByDocx = new Map(parsed.blocks.map((b) => [b.docxIndex, b]))
        let cursor = 0
        for (const block of parsed.blocks) {
          if (block.hidden) continue
          const runs = (block.runs ?? []).map((r) => ({
            text: r.text ?? '',
            ...(r.font ? { fontFamily: r.font } : {}),
            sizeHalfPoints: r.sizeHalfPoints ?? 24,
          }))
          const result = computeLineMetrics({
            runs,
            availWidthPx: contentWidthPx,
            lineRule: parsed.docDefaults?.lineRule,
            lineRawTwips: parsed.docDefaults?.lineRawTwips,
            spaceAfter: parsed.docDefaults?.spaceAfterTwips ?? 160,
            docGrid,
            defaultFontSizePt: 12,
            metrics,
          })
          const bands: Array<{ offset: number; height: number }> = []
          for (const run of block.runs ?? []) {
            if (run.noteRef?.kind !== 'footnote') continue
            const fn = footnoteById.get(run.noteRef.id)
            if (fn) bands.push({ offset: Infinity, height: heightOf(fn) })
          }
          const footnoteExtra = bands.reduce((s, b) => s + b.height, 0)
          const totalH = result.totalHeight + footnoteExtra
          blocks.push({
            top: cursor,
            height: Math.max(totalH, 1),
            lineBoxes: result.lineBoxes,
            spaceBeforePx: result.spaceBeforePx,
            spaceAfterPx: result.spaceAfterPx,
            ...(footnoteExtra > 0 ? { footnoteExtraPx: footnoteExtra, noteBands: bands } : {}),
            keepNext: block.styleId === 'Heading1' || undefined,
            docxIndex: block.docxIndex ?? undefined,
            section: 0,
          })
          cursor += Math.max(totalH, 1)
        }
        const slices = computeSectionedSlicesF2(blocks, geoms, cursor)
        return slices.map((slice) => {
          const b = blocks.find((x) => x.top >= slice.start - 1 && x.docxIndex !== undefined)
          const src = b ? blockByDocx.get(b.docxIndex!) : undefined
          return (src?.runs ?? []).map((r) => r.text).join('')
        })
      }

      // the defect model demonstrably misses Word's page 1 break (its page 2
      // starts past paragraph 14) — the guard that keeps the spacing term in
      const bareStarts = paginateWith(bareHeightOf)
      const fixedStarts = paginateWith(reservedHeightOf)

      if (baseline) {
        // page 1 holds title + paragraphs 1-13 like Word: page 2 starts at paragraph 14
        expect(fixedStarts).toHaveLength(baseline.pages)
        expect(fixedStarts[1].startsWith('第14段正文')).toBe(true)
        for (let i = 1; i < baseline.pages; i++) {
          const baseKey = baseline.pageStarts[i].replace(/\s+/g, '')
          expect(fixedStarts[i].replace(/\s+/g, '').startsWith(baseKey.slice(0, 8))).toBe(true)
        }
        // and the bare-row model this fix replaced could not match it
        expect(bareStarts[1].startsWith('第14段正文')).toBe(false)
      } else {
        // baseline not collected in this environment: pin the break directly
        expect(fixedStarts[1].startsWith('第14段正文')).toBe(true)
        expect(bareStarts[1].startsWith('第14段正文')).toBe(false)
      }
    },
  )

  it('noteReservedHeightPx adds the resolved before/after and tolerates no style', () => {
    expect(noteReservedHeightPx(20)).toBe(20)
    expect(
      noteReservedHeightPx(20, { sizeHalfPoints: 18, spaceBeforeTwips: 40, spaceAfterTwips: 80 }),
    ).toBeCloseTo(20 + ((40 + 80) * 96) / 1440, 6)
  })
})

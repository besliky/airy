/**
 * BUG-1708 regression: sections that share (link) the same header/footer parts
 * must stay separate sections and paginate one page per section, and a document
 * whose body-level trailing sectPr is missing must still give the blocks after
 * the last paragraph-level sectPr their own (implicit) final section — Word
 * treats the missing final properties as defaults, not as "absorbed into the
 * previous section".
 *
 * Page counts use the headless F2 slicer with engine-style line measurement
 * (the same harness as pagination-parity.test.ts).
 */
import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, type ParsedDoc } from '@airy-office/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  computeSectionedSlicesF2,
  effectiveHfRefs,
  sectionColGeom,
  sectionGeoms,
  type BlockBox,
} from '../src/renderer/pagination'
import { computeLineMetrics } from '../src/renderer/line-metrics'
import { loBaselineMetrics } from './helpers/lo-fonts'
import JSZip from 'jszip'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

const SHARED_REFS =
  '<w:headerReference w:type="default" r:id="rIdH1"/>' +
  '<w:headerReference w:type="even" r:id="rIdH2"/>' +
  '<w:footerReference w:type="default" r:id="rIdF1"/>'

const breakWithSharedRefs = (opts: { landscape?: boolean; titlePg?: boolean } = {}) => {
  const size = opts.landscape
    ? '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
    : '<w:pgSz w:w="12240" w:h="15840"/>'
  return (
    `<w:p><w:pPr><w:sectPr>${SHARED_REFS}` +
    (opts.titlePg ? '<w:titlePg/>' : '') +
    size +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '</w:sectPr></w:pPr></w:p>'
  )
}

const EXTRA_RELS =
  '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
  '<Relationship Id="rIdH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>' +
  '<Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'

const HF_PART_XML = (tag: 'w:hdr' | 'w:ftr', text: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></${tag}>`

async function buildSharedRefDoc(bodyXml: string): Promise<Uint8Array> {
  return buildDocx({
    bodyXml,
    sectPrExtra: SHARED_REFS,
    extraRels: EXTRA_RELS,
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: HF_PART_XML('w:hdr', 'SHARED-HEADER-ONE'),
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'word/header2.xml',
        xml: HF_PART_XML('w:hdr', 'SHARED-HEADER-EVEN'),
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'word/footer1.xml',
        xml: HF_PART_XML('w:ftr', 'SHARED-FOOTER-ONE'),
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      },
    ],
  })
}

async function stripTrailingSectPr(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const doc = await zip.file('word/document.xml')!.async('string')
  const start = doc.lastIndexOf('<w:sectPr')
  const end = doc.lastIndexOf('</w:sectPr></w:body>')
  zip.file('word/document.xml', doc.slice(0, start) + doc.slice(end + '</w:sectPr>'.length))
  return zip.generateAsync({ type: 'uint8array' })
}

/** headless page count: engine-model line boxes + the F2 sectioned slicer */
async function paginate(parsed: ParsedDoc): Promise<{ pages: number; sections: unknown[] }> {
  const sections = readSections(parsed)
  const metrics = loBaselineMetrics()
  const geoms = sectionGeoms(sections)
  const widths = sections.map((s) => sectionColGeom(s).colWidthPx)
  const blocks: BlockBox[] = []
  let cursor = 0
  let cur = 0
  for (const block of parsed.blocks) {
    if (block.hidden) continue
    if (
      block.type === 'passthrough' &&
      block.originalXml?.includes('<w:sectPr') &&
      !block.originalXml?.includes('<w:t>')
    )
      continue
    const idx = block.docxIndex ?? 0
    let si = cur
    for (let i = 0; i < sections.length; i++) {
      if (idx <= sections[i].lastBlockIndex) {
        si = i
        break
      }
    }
    cur = si
    const runs = block.runs ?? []
    const r = computeLineMetrics({
      runs: runs.map((x) => ({ text: x.text ?? '' })),
      availWidthPx: widths[si],
      defaultFontSizePt: 12,
      metrics,
      isEmpty: !runs.some((x) => x.text),
    })
    blocks.push({
      top: cursor,
      height: r.totalHeight,
      lineBoxes: r.lineBoxes,
      docxIndex: block.docxIndex ?? undefined,
      section: si,
    })
    cursor += r.totalHeight
  }
  const slices = computeSectionedSlicesF2(blocks, geoms, cursor)
  return {
    pages: slices.length,
    sections: slices.map((s) => s.section),
  }
}

describe('BUG-1708: shared header/footer parts between sections', () => {
  it('3 sections with identical refs stay 3 sections and paginate 3 pages', async () => {
    const parsed = await parseDocx(
      await buildSharedRefDoc(
        P('S1 line one') +
          P('S1 line two') +
          breakWithSharedRefs({ titlePg: true }) +
          P('S2 line one') +
          P('S2 line two') +
          breakWithSharedRefs({ landscape: true }) +
          P('S3 line one') +
          P('S3 line two'),
      ),
    )
    const sections = readSections(parsed)
    expect(sections.length).toBe(3)
    expect(sections.map((s) => s.settings.orientation)).toEqual([
      'portrait',
      'landscape',
      'portrait',
    ])
    for (const s of sections) {
      expect(s.headerRefs.default).toBe('rIdH1')
      expect(s.headerRefs.even).toBe('rIdH2')
      expect(s.footerRefs.default).toBe('rIdF1')
    }
    const result = await paginate(parsed)
    expect(result.pages).toBe(3)
    expect(result.sections).toEqual([0, 1, 2])
  })

  it('absent refs resolve by inheritance (linked to previous), shared refs stay verbatim', () => {
    const info = effectiveHfRefs([
      {
        headerRefs: { default: 'rIdH1', even: 'rIdH2' },
        footerRefs: { default: 'rIdF1' },
      },
      // section 2 links by absence: nothing of its own
      { headerRefs: {}, footerRefs: {} },
      // section 3 overrides only the default header; even + footer inherit
      { headerRefs: { default: 'rIdH9' }, footerRefs: {} },
    ] as never)
    expect(info[0].header).toEqual({ default: 'rIdH1', even: 'rIdH2' })
    expect(info[0].footer).toEqual({ default: 'rIdF1' })
    expect(info[1].header).toEqual({ default: 'rIdH1', even: 'rIdH2' })
    expect(info[1].footer).toEqual({ default: 'rIdF1' })
    expect(info[2].header).toEqual({ default: 'rIdH9', even: 'rIdH2' })
    expect(info[2].footer).toEqual({ default: 'rIdF1' })
  })

  it('missing trailing sectPr: the tail paragraphs form the implicit final section (3 pages)', async () => {
    const source = await buildSharedRefDoc(
      P('S1 first') +
        P('S1 second') +
        breakWithSharedRefs({ titlePg: true }) +
        P('S2 first') +
        P('S2 second') +
        breakWithSharedRefs({ landscape: true }) +
        P('S3 closing'),
    )
    const parsed = await parseDocx(await stripTrailingSectPr(source))
    const sections = readSections(parsed)
    // sectPr 1 closes the portrait section, sectPr 2 the landscape one, and the
    // trailing paragraph (no body-level sectPr) gets the implicit final section
    expect(sections.length).toBe(3)
    expect(sections.map((s) => s.settings.orientation)).toEqual([
      'portrait',
      'landscape',
      'portrait',
    ])
    const tail = sections[2]
    expect(tail.firstBlockIndex).toBe(tail.lastBlockIndex)
    expect(tail.headerRefs).toEqual({})
    const result = await paginate(parsed)
    expect(result.pages).toBe(3)
    expect(result.sections).toEqual([0, 1, 2])
  })
})

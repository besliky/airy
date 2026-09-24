import { execFileSync } from 'node:child_process'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  StandardFonts,
  degrees,
} from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { pageRotation, repairBrokenPageRotations, snapRotation } from '../src/main/page-rotation'
import {
  applySaveRequest,
  cropPagesBytes,
  insertBlankPageBytes,
  setPageSizeBytes,
  splitPagesBytes,
} from '../src/main/save-pdf'
import type { MarkupInput, SavePdfRequest } from '../src/shared/ipc'

const ROTATE_KEY = PDFName.of('Rotate')

/** The page's own /Rotate value (as opposed to the tolerant inherited read) */
const ownRotate = (doc: PDFDocument, pageIndex: number): number =>
  (doc.getPage(pageIndex).node.get(ROTATE_KEY) as PDFNumber).asNumber()

/** One-page PDF with text and a malformed `/Rotate /90` (a name where the spec wants an integer) */
async function makeBrokenPdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 200])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('BROKEN ROTATE PAGE', { x: 20, y: 150, font, size: 12 })
  page.node.set(ROTATE_KEY, PDFName.of('90'))
  return doc.save({ useObjectStreams: false })
}

async function makePdfWithRotation(angle: number | null): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 200])
  if (angle !== null) page.setRotation(degrees(angle))
  return doc.save({ useObjectStreams: false })
}

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

const underlineMarkup = (): MarkupInput => ({
  pageIndex: 0,
  type: 'underline',
  color: [0.17, 0.4, 1],
  quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
})

/** Text of the /AP /N appearance stream of the i-th annotation on the page */
function annotApText(doc: PDFDocument, pageIndex: number, index: number): string {
  const annots = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  const dict = annots!.lookup(index, PDFDict)
  const nRef = dict.lookup(PDFName.of('AP'), PDFDict).get(PDFName.of('N'))
  return new TextDecoder().decode((doc.context.lookup(nRef) as PDFRawStream).contents)
}

const hasPoppler = ((): boolean => {
  try {
    execFileSync('which', ['pdfinfo'], { stdio: 'ignore' })
    execFileSync('which', ['pdftotext'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
})()

describe('snapRotation', () => {
  it('snaps arbitrary values to the nearest valid multiple of 90', () => {
    expect(snapRotation(0)).toBe(0)
    expect(snapRotation(45)).toBe(90)
    expect(snapRotation(90)).toBe(90)
    expect(snapRotation(137)).toBe(180)
    expect(snapRotation(225)).toBe(270)
    expect(snapRotation(400)).toBe(0)
    expect(snapRotation(-90)).toBe(270)
  })

  it('reads non-finite values as 0', () => {
    expect(snapRotation(Number.NaN)).toBe(0)
    expect(snapRotation(Number.POSITIVE_INFINITY)).toBe(0)
  })
})

describe('pageRotation', () => {
  it('reads valid numeric rotations unchanged', async () => {
    for (const angle of [0, 90, 180, 270]) {
      const doc = await PDFDocument.load(await makePdfWithRotation(angle))
      expect(pageRotation(doc.getPage(0))).toBe(angle)
    }
  })

  it('reads a missing /Rotate as 0', async () => {
    const doc = await PDFDocument.load(await makePdfWithRotation(null))
    expect(pageRotation(doc.getPage(0))).toBe(0)
  })

  it('reads a malformed /Rotate name as 0 instead of throwing', async () => {
    const doc = await PDFDocument.load(await makeBrokenPdf())
    // Sanity: the raw pdf-lib call this getter replaces does throw here
    expect(() => doc.getPage(0).getRotation()).toThrow()
    expect(pageRotation(doc.getPage(0))).toBe(0)
  })

  it('reads a numeric /Rotate that is not a multiple of 90 as the nearest valid one', async () => {
    const doc = await PDFDocument.load(await makePdfWithRotation(null))
    doc.getPage(0).node.set(ROTATE_KEY, PDFNumber.of(45))
    expect(pageRotation(doc.getPage(0))).toBe(90)
  })

  it('resolves an indirect-ref /Rotate', async () => {
    const doc = await PDFDocument.load(await makePdfWithRotation(null))
    doc.getPage(0).node.set(ROTATE_KEY, doc.context.register(PDFNumber.of(180)))
    expect(pageRotation(doc.getPage(0))).toBe(180)
  })

  it('follows the inheritable chain and tolerates a malformed inherited value', async () => {
    const doc = await PDFDocument.load(await makePdfWithRotation(null))
    const pagesNode = doc.context.lookupMaybe(doc.catalog.get(PDFName.of('Pages')), PDFDict)
    pagesNode!.set(ROTATE_KEY, PDFName.of('270'))
    expect(pageRotation(doc.getPage(0))).toBe(0)
  })
})

describe('repairBrokenPageRotations', () => {
  it('replaces malformed entries with numeric 0 and keeps valid ones', async () => {
    const doc = await PDFDocument.load(await makeBrokenPdf())
    doc.addPage([300, 200]).setRotation(degrees(90))
    repairBrokenPageRotations(doc)
    expect(doc.getPage(0).node.get(ROTATE_KEY)).toBeInstanceOf(PDFNumber)
    expect(ownRotate(doc, 0)).toBe(0)
    expect(ownRotate(doc, 1)).toBe(90)
  })

  it('overrides a malformed value inherited from the page tree', async () => {
    const doc = await PDFDocument.load(await makePdfWithRotation(null))
    const pagesNode = doc.context.lookupMaybe(doc.catalog.get(PDFName.of('Pages')), PDFDict)
    pagesNode!.set(ROTATE_KEY, PDFName.of('270'))
    repairBrokenPageRotations(doc)
    expect(doc.getPage(0).node.get(ROTATE_KEY)).toBeInstanceOf(PDFNumber)
    expect(ownRotate(doc, 0)).toBe(0)
  })
})

describe('applySaveRequest on a file with a malformed /Rotate (BUG-1668)', () => {
  it('annotates, applies a rotation delta and saves without failing', async () => {
    const broken = await makeBrokenPdf()
    const saved = await applySaveRequest(
      broken,
      request({
        markups: [
          {
            pageIndex: 0,
            type: 'highlight',
            color: [1, 0.87, 0.35],
            quads: [[20, 155, 140, 155, 20, 145, 140, 145]],
          },
        ],
        drawings: [{ kind: 'note', pageIndex: 0, color: [1, 0, 0], at: [200, 170], contents: 'x' }],
        // The malformed name reads as 0, so a +90 delta lands on exactly 90
        rotations: [{ pageIndex: 0, delta: 90 }],
      }),
    )
    const out = await PDFDocument.load(saved.bytes)
    expect(out.getPageCount()).toBe(1)
    // The malformed name is repaired, so the +90 delta lands on a numeric 90
    expect(ownRotate(out, 0)).toBe(90)
    const annots = out.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    expect(annots!.size()).toBe(2)
  })

  it('produces output that external viewers read', async () => {
    const saved = await applySaveRequest(
      await makeBrokenPdf(),
      request({ markups: [underlineMarkup()] }),
    )
    // pdf-lib structure: readable, annot present, no malformed /Rotate left
    const out = await PDFDocument.load(saved.bytes)
    expect(out.getPageCount()).toBe(1)
    const annots = out.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    expect(annots!.size()).toBe(1)
    expect(new TextDecoder().decode(saved.bytes)).not.toContain('/Rotate /90')

    // pdf.js (the renderer engine) opens it
    const loadingTask = getDocument({ data: saved.bytes.slice() })
    try {
      const pdfJsDoc = await loadingTask.promise
      const page = await pdfJsDoc.getPage(1)
      expect(pdfJsDoc.numPages).toBe(1)
      expect(page.rotate).toBe(0)
    } finally {
      await loadingTask.destroy()
    }

    // poppler (when installed) parses the file and still finds the page text
    if (hasPoppler) {
      const dir = mkdtempSync(join(tmpdir(), 'bug1668-'))
      const path = join(dir, 'out.pdf')
      writeFileSync(path, saved.bytes)
      const info = execFileSync('pdfinfo', [path], { encoding: 'utf8' })
      expect(info).toContain('Pages:')
      const text = execFileSync('pdftotext', [path, '-'], { encoding: 'utf8' })
      expect(text).toContain('BROKEN ROTATE PAGE')
    }
  })
})

describe('valid rotations keep their save behavior', () => {
  // The underline appearance segment for each final orientation (save-pdf's
  // markupAppearance branch table) must be unchanged by the tolerance work
  it.each([
    [0, '10 88.96 m 60 88.96 l S'],
    [90, '56 88 m 56 100 l S'],
    [180, '10 99.04 m 60 99.04 l S'],
    [270, '14 88 m 14 100 l S'],
  ])('draws the underline appearance for /Rotate %i', async (angle, segment) => {
    const saved = await applySaveRequest(
      await makePdfWithRotation(angle),
      request({ markups: [underlineMarkup()] }),
    )
    const out = await PDFDocument.load(saved.bytes)
    expect(ownRotate(out, 0)).toBe(angle)
    expect(annotApText(out, 0, 0)).toContain(segment)
  })
})

describe('page operations tolerate a malformed /Rotate', () => {
  it('setPageSizeBytes treats the page as unrotated', async () => {
    const out = await PDFDocument.load(await setPageSizeBytes(await makeBrokenPdf(), 595, 842))
    expect(out.getPage(0).getWidth()).toBe(595)
    expect(out.getPage(0).getHeight()).toBe(842)
  })

  it('splitPagesBytes splits the page as displayed', async () => {
    const out = await PDFDocument.load(await splitPagesBytes(await makeBrokenPdf(), 2))
    expect(out.getPageCount()).toBe(2)
  })

  it('cropPagesBytes maps the crop rect through rotation 0', async () => {
    const out = await PDFDocument.load(
      await cropPagesBytes(await makeBrokenPdf(), [0], { l: 0.1, t: 0.2, r: 0.6, b: 0.7 }),
    )
    expect(out.getPage(0).getCropBox().x).toBeCloseTo(30)
    expect(out.getPage(0).getCropBox().y).toBeCloseTo(60)
    expect(out.getPage(0).getCropBox().width).toBeCloseTo(150)
    expect(out.getPage(0).getCropBox().height).toBeCloseTo(100)
  })

  it('insertBlankPageBytes inserts after the broken page with rotation 0', async () => {
    const out = await PDFDocument.load(await insertBlankPageBytes(await makeBrokenPdf(), 0))
    expect(out.getPageCount()).toBe(2)
    expect(ownRotate(out, 1)).toBe(0)
  })
})

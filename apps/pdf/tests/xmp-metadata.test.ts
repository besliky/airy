import { PDFArray, PDFDocument, PDFName, PDFRawStream, PDFRef, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applySaveRequest } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

/**
 * Regression tests for BUG-1667: the document XMP metadata stream (catalog /Metadata,
 * /Subtype /XML — dc:title, xmp:CreateDate, pdf:Producer, ...) must survive every save.
 * pdf-lib's load/save keeps it, but the pdfium rewrite stages (text edits, text inserts,
 * image edits, annotation deletes) drop it via pdfium's SaveAsCopy; the save pipeline
 * re-attaches the byte-identical source stream.
 */

const XMP_TITLE = 'XMP-PROBE-TITLE-24'

const XMP = `<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title><rdf:Alt><rdf:li xml:lang="x-default">${XMP_TITLE}</rdf:li></rdf:Alt></dc:title></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>`

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

/** One-page PDF with a single Helvetica text run at a known position */
async function textPdf(text: string): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText(text, { x: 50, y: 700, size: 14, font })
  return doc.save({ useObjectStreams: false })
}

const RUN_RECT = (): [number, number, number, number] => {
  // Helvetica 14pt width of "Alpha zero run" is stable enough for the matcher's tolerance
  return [45, 694, 145, 718]
}

/**
 * Turn pdf-lib output into a single-generation source whose catalog carries
 * /Metadata -> a raw XML stream (/Type /Metadata /Subtype /XML), rebuilding the
 * xref table so regular readers see the metadata too.
 */
function injectXmp(bytes: Uint8Array): Uint8Array {
  const s = Buffer.from(bytes).toString('latin1')
  const catNum = Number(s.match(/(\d+) 0 obj\s*<<\s*\/Type \/Catalog/)![1])
  const metaNum = Math.max(...[...s.matchAll(/(\d+) 0 obj/g)].map((m) => Number(m[1]))) + 1
  const metaObj = `${metaNum} 0 obj\n<< /Type /Metadata /Subtype /XML /Length ${XMP.length} >>\nstream\n${XMP}\nendstream\nendobj\n`
  const catStart = s.indexOf(`${catNum} 0 obj`)
  const catEnd = s.indexOf('endobj', catStart)
  const newCatalog = s
    .slice(catStart, catEnd)
    .replace('/Type /Catalog\n/Pages', `/Type /Catalog\n/Metadata ${metaNum} 0 R\n/Pages`)
  const body = s.slice(0, catStart) + newCatalog + s.slice(catEnd) + metaObj
  // Rebuild the xref table: the injected object shifted nothing, but it was appended
  // after the original xref, so the table must be rewritten to cover it
  const nl = body.indexOf('\n') + 1
  const header = body.slice(0, nl)
  const objects = body.slice(nl)
  const offsets = new Map<number, number>()
  for (const m of objects.matchAll(/(\d+) 0 obj/g)) offsets.set(Number(m[1]), m.index!)
  const maxNum = Math.max(...offsets.keys())
  let xref = `xref\n0 ${maxNum + 1}\n0000000000 65535 f \n`
  for (let i = 1; i <= maxNum; i++) {
    xref += offsets.has(i)
      ? `${String(offsets.get(i)).padStart(10, '0')} 00000 n \n`
      : '0000000000 65535 f \n'
  }
  const trailer = `trailer\n<< /Size ${maxNum + 1} /Root ${catNum} 0 R >>\nstartxref\n${header.length + objects.length}\n%%EOF\n`
  return new Uint8Array(Buffer.from(header + objects + xref + trailer, 'latin1'))
}

/**
 * A source whose XMP was added by a legal incremental update: the first generation has
 * a plain catalog, the appended second generation carries the XMP stream and a new
 * catalog generation with /Metadata, selected by the updated trailer /Root.
 */
function appendXmpGeneration(base: Uint8Array): Uint8Array {
  const s = Buffer.from(base).toString('latin1')
  const objNums = [...s.matchAll(/(\d+) 0 obj/g)].map((m) => Number(m[1]))
  const maxNum = Math.max(...objNums)
  const catNum = Number(s.match(/\/Root (\d+) \d+ R/)![1])
  const pagesRef = s.match(new RegExp(`${catNum} 0 obj[\\s\\S]*?/Pages (\\d+ 0 R)`))![1]
  const xmpNum = maxNum + 1
  const newCatNum = maxNum + 2
  const xmpObj = `${xmpNum} 0 obj\n<< /Type /Metadata /Subtype /XML /Length ${XMP.length} >>\nstream\n${XMP}\nendstream\nendobj\n`
  const catObj = `${newCatNum} 0 obj\n<< /Type /Catalog /Pages ${pagesRef} /Metadata ${xmpNum} 0 R >>\nendobj\n`
  const sectionOffset = base.length
  const catOffset = sectionOffset + xmpObj.length
  const prevStart = Number(
    s
      .slice(s.lastIndexOf('startxref') + 9)
      .trim()
      .split(/\s+/)[0],
  )
  const xref = `xref\n${xmpNum} 2\n${String(sectionOffset).padStart(10, '0')} 00000 n \n${String(catOffset).padStart(10, '0')} 00000 n \ntrailer\n<< /Size ${maxNum + 3} /Root ${newCatNum} 0 R /Prev ${prevStart} >>\nstartxref\n${sectionOffset + xmpObj.length + catObj.length}\n%%EOF\n`
  return new Uint8Array(Buffer.from(s + section(xmpObj + catObj + xref), 'latin1'))
}

const section = (body: string) => `\n%incremental-xmp-update\n${body}`

/** XMP stream contents of a document, byte-exact as stored, or undefined when absent */
async function xmpOf(bytes: Uint8Array): Promise<{ xml: string; subtype: string } | undefined> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false, ignoreEncryption: true })
  const entry = doc.catalog.get(PDFName.of('Metadata'))
  if (!(entry instanceof PDFRef)) return undefined
  const stream = doc.context.lookupMaybe(entry, PDFRawStream)
  if (!stream) return undefined
  const subtype = stream.dict.lookup(PDFName.of('Subtype'), PDFName)
  return { xml: Buffer.from(stream.contents).toString('latin1'), subtype: String(subtype) }
}

const expectByteIdenticalXmp = async (bytes: Uint8Array) => {
  const xmp = await xmpOf(bytes)
  expect(xmp).toBeDefined()
  expect(xmp!.xml).toBe(XMP)
  expect(xmp!.subtype).toBe('/XML')
  expect(Buffer.from(bytes).toString('latin1')).toContain(XMP_TITLE)
}

describe('XMP metadata survives save (BUG-1667)', () => {
  it('keeps the XMP stream byte-for-byte on a save without content edits', async () => {
    const { bytes } = await applySaveRequest(
      await injectXmp(await textPdf('Alpha zero run')),
      request(),
    )
    await expectByteIdenticalXmp(bytes)
  })

  it('keeps the XMP stream byte-for-byte when a text edit runs the pdfium rewrite', async () => {
    // Regression: pdfium's SaveAsCopy (text-edit stage) drops the catalog /Metadata
    const { bytes, skippedTextEdits } = await applySaveRequest(
      await injectXmp(await textPdf('Alpha zero run')),
      request({
        textEdits: [
          {
            pageIndex: 0,
            rect: RUN_RECT(),
            oldText: 'Alpha zero run',
            newText: 'Beta one run',
            fontSize: 14,
          },
        ],
      }),
    )
    expect(skippedTextEdits).toEqual([])
    await expectByteIdenticalXmp(bytes)
    // The restored stream must be reachable through the rebuilt catalog: a second
    // engine (pdf.js) reads the XMP back from the saved bytes
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const loadingTask = getDocument({ data: bytes.slice() })
    try {
      const pdfJsDoc = await loadingTask.promise
      expect(pdfJsDoc.numPages).toBe(1)
      const metadata = await pdfJsDoc.getMetadata()
      expect(metadata.metadata?.getRaw()).toContain(XMP_TITLE)
    } finally {
      await loadingTask.destroy()
    }
  })

  it('keeps the XMP stream byte-for-byte when a text insert runs the pdfium rewrite', async () => {
    const { bytes } = await applySaveRequest(
      await injectXmp(await textPdf('Alpha zero run')),
      request({
        textInserts: [
          {
            pageIndex: 0,
            origin: [50, 650],
            text: 'Inserted line',
            fontSize: 14,
            color: [0, 0, 0],
          },
        ],
      }),
    )
    await expectByteIdenticalXmp(bytes)
  })

  it('keeps the XMP stream byte-for-byte when an annotation delete runs the pdfium rewrite', async () => {
    const source = await injectXmp(await textPdf('Alpha zero run'))
    // Write a highlight through the real save path (this save preserves the XMP too)
    const annotated = await applySaveRequest(
      source,
      request({
        markups: [
          {
            pageIndex: 0,
            type: 'highlight',
            color: [1, 0.87, 0.35],
            quads: [[60, 700, 200, 700, 60, 685, 200, 685]],
          },
        ],
      }),
    ).then((r) => r.bytes)
    const doc = await PDFDocument.load(annotated)
    const annots = doc.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    const ref = annots!.get(0) as PDFRef
    const { bytes } = await applySaveRequest(
      annotated,
      request({
        annotDeletes: [
          {
            pageIndex: 0,
            objNum: ref.objectNumber,
            subtype: 'highlight',
            rect: [60, 685, 200, 700],
          },
        ],
      }),
    )
    await expectByteIdenticalXmp(bytes)
  })

  it('keeps XMP added by an incremental update (second catalog generation)', async () => {
    const source = appendXmpGeneration(await textPdf('Alpha zero run'))
    // The fixture itself must expose the XMP before the save
    await expectByteIdenticalXmp(source)
    const { bytes } = await applySaveRequest(source, request())
    await expectByteIdenticalXmp(bytes)
  })

  it('treats explicit metadata edits as Info-only: the XMP keeps its authored bytes', async () => {
    // Conscious choice: the Properties dialog edits the Info dictionary only; the
    // original XMP is never rewritten, so the two views may diverge after an edit
    const { bytes } = await applySaveRequest(
      await injectXmp(await textPdf('Alpha zero run')),
      request({ metadata: { title: 'Info Only Title' } }),
    )
    const out = await PDFDocument.load(bytes)
    expect(out.getTitle()).toBe('Info Only Title')
    await expectByteIdenticalXmp(bytes)
  })

  it('adds no /Metadata key when the source document never had XMP', async () => {
    const { bytes } = await applySaveRequest(
      await textPdf('Alpha zero run'),
      request({
        textEdits: [
          {
            pageIndex: 0,
            rect: RUN_RECT(),
            oldText: 'Alpha zero run',
            newText: 'Beta one run',
            fontSize: 14,
          },
        ],
      }),
    )
    expect(await xmpOf(bytes)).toBeUndefined()
    expect(Buffer.from(bytes).toString('latin1')).not.toContain('/Metadata')
  })
})

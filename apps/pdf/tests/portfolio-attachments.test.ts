import { describe, expect, it } from 'vitest'
import { PDFDocument, PDFName, PDFDict } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applySaveRequest } from '../src/main/save-pdf'

/**
 * UX-1734: a PDF portfolio (/Collection + /Names/EmbeddedFiles children) shows
 * only its cover page in the viewer. Two things must hold: the children are
 * enumerable by pdf.js in the renderer (the Attachments panel's data source),
 * and they survive the save pipeline — an edit to the cover must not drop the
 * embedded files.
 */

async function makePortfolio(): Promise<{
  bytes: Uint8Array
  childA: Uint8Array
  childB: Uint8Array
}> {
  const doc = await PDFDocument.create()
  doc.addPage([200, 300])
  const childA = await PDFDocument.create()
  childA.addPage([100, 100])
  const bytesA = await childA.save()
  const childB = await PDFDocument.create()
  childB.addPage([120, 140])
  const bytesB = await childB.save()
  await doc.attach(bytesA, 'child-a.pdf', { mimeType: 'application/pdf' })
  await doc.attach(bytesB, 'child-b.pdf', { mimeType: 'application/pdf' })
  // the canonical portfolio shape: a /Collection dict in the catalog
  doc.catalog.set(PDFName.of('Collection'), doc.context.obj({ Type: 'Collection' }))
  return { bytes: await doc.save({ useObjectStreams: false }), childA: bytesA, childB: bytesB }
}

interface AttachmentsMap extends Map<string, { filename: string }> {
  get(id: string): { filename: string } | undefined
}

async function readAttachments(bytes: Uint8Array): Promise<Map<string, { filename: string }>> {
  const doc = await getDocument({ data: bytes.slice() }).promise
  const raw = (await doc.getAttachments()) as AttachmentsMap | null
  const out = new Map<string, { filename: string }>()
  if (raw) for (const [id, a] of raw.entries()) out.set(id, { filename: a.filename })
  await doc.loadingTask.destroy()
  return out
}

describe('PDF portfolio attachments (UX-1734)', () => {
  it('enumerates the embedded children through pdf.js with readable content', async () => {
    const { bytes, childA, childB } = await makePortfolio()
    const doc = await getDocument({ data: bytes.slice() }).promise
    const attachments = await doc.getAttachments()
    expect(attachments).not.toBeNull()
    const byName = new Map([...attachments!.entries()].map(([id, a]) => [a.filename, id]))
    expect([...byName.keys()].sort()).toEqual(['child-a.pdf', 'child-b.pdf'])
    // the ids are the keys getAttachmentContent resolves; the bytes must be intact
    const contentA = await doc.getAttachmentContent(byName.get('child-a.pdf')!)
    const contentB = await doc.getAttachmentContent(byName.get('child-b.pdf')!)
    expect(Buffer.from(contentA!)).toEqual(Buffer.from(childA))
    expect(Buffer.from(contentB!)).toEqual(Buffer.from(childB))
    await doc.loadingTask.destroy()
  })

  it('children survive a save that edits the cover page', async () => {
    const { bytes, childA, childB } = await makePortfolio()
    const result = await applySaveRequest(bytes, {
      path: '/tmp/portfolio.pdf',
      markups: [
        {
          pageIndex: 0,
          type: 'highlight',
          color: [1, 0.9, 0.2],
          quads: [[10, 280, 190, 290, 10, 270, 190, 280]],
        },
      ],
      drawings: [],
      formValues: [],
      stamps: [],
    })
    expect(result.skippedTextEdits).toEqual([])

    // the attachments enumeration is unchanged after the rewrite
    const after = await readAttachments(result.bytes)
    expect([...after.keys()].sort()).toEqual(['child-a.pdf', 'child-b.pdf'])

    // byte identity: the same embedded payloads come back out
    const doc = await getDocument({ data: result.bytes.slice() }).promise
    const attachments = await doc.getAttachments()
    const byName = new Map([...attachments!.entries()].map(([id, a]) => [a.filename, id]))
    expect(Buffer.from((await doc.getAttachmentContent(byName.get('child-a.pdf')!))!)).toEqual(
      Buffer.from(childA),
    )
    expect(Buffer.from((await doc.getAttachmentContent(byName.get('child-b.pdf')!))!)).toEqual(
      Buffer.from(childB),
    )
    await doc.loadingTask.destroy()
  })

  it('keeps the portfolio /Collection marker through a save', async () => {
    const { bytes } = await makePortfolio()
    const result = await applySaveRequest(bytes, {
      path: '/tmp/portfolio.pdf',
      rotations: [{ pageIndex: 0, delta: 90 }],
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
    })
    const doc = await PDFDocument.load(result.bytes)
    expect(doc.catalog.lookupMaybe(PDFName.of('Collection'), PDFDict)).not.toBeUndefined()
  })
})

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRef, StandardFonts } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { insertPdfBytes, InsertSourceLoadError, sourcePageShapes } from '../src/main/save-pdf'

/** Labeled pages: one text marker per page, distinct sizes, so order and
    content are both observable in the merged output */
async function makeLabeledPdf(
  labels: string[],
  size: [number, number] = [200, 300],
): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (const label of labels) {
    doc.addPage(size).drawText(label, { x: 20, y: size[1] - 50, size: 24, font })
  }
  return doc.save({ useObjectStreams: false })
}

/** Extracted text per page (pdf.js) — proves transported content streams */
async function pageTexts(bytes: Uint8Array): Promise<string[]> {
  const task = getDocument({ data: bytes.slice() })
  try {
    const doc = await task.promise
    const out: string[] = []
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i)
      const content = await page.getTextContent()
      out.push(content.items.map((item) => ('str' in item ? item.str : '')).join(''))
    }
    return out
  } finally {
    await task.destroy()
  }
}

/** Hand-crafted trailer-level encryption stub: parses as a PDF but pdf-lib
    marks it encrypted (real encryption is not needed for the refusal path) */
const ENCRYPTED_PDF = Buffer.from(
  '%PDF-1.4\n' +
    '1 0 obj\n<</Type/Catalog/Pages 2 0 R>>\nendobj\n' +
    '2 0 obj\n<</Type/Pages/Kids[3 0 R]/Count 1>>\nendobj\n' +
    '3 0 obj\n<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 300]>>\nendobj\n' +
    '4 0 obj\n<</Filter/Standard/V/1/R 2/O(s)/U(s)>>\nendobj\n' +
    'xref\n0 5\n0000000000 65535 f \n0000000009 00000 n \n0000000060 00000 n \n' +
    '0000000111 00000 n \n0000000174 00000 n \n' +
    'trailer\n<</Size 5/Root 1 0 R/Encrypt 4 0 R>>\nstartxref\n231\n%%EOF',
  'latin1',
)

function pageAnnots(doc: PDFDocument, pageIndex: number): PDFDict[] {
  const annots = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  return Array.from({ length: annots.size() }, (_, i) => annots.lookup(i, PDFDict))
}

const subtypeOf = (annot: PDFDict) => annot.lookup(PDFName.of('Subtype'), PDFName).decodeText()

describe('sourcePageShapes', () => {
  it('reports page count and per-page sizes for the pick preview', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([200, 300])
    doc.addPage([612, 792])
    const shapes = await sourcePageShapes(await doc.save({ useObjectStreams: false }))
    expect(shapes).toEqual([
      { width: 200, height: 300 },
      { width: 612, height: 792 },
    ])
  })

  it('refuses a password-protected source with the encrypted kind', async () => {
    const err = await sourcePageShapes(new Uint8Array(ENCRYPTED_PDF)).catch((e: unknown) => e)
    expect(err).toBeInstanceOf(InsertSourceLoadError)
    expect((err as InsertSourceLoadError).kind).toBe('encrypted')
  })

  it('refuses a non-PDF file with the invalid kind', async () => {
    const err = await sourcePageShapes(new Uint8Array(Buffer.from('not a pdf'))).catch(
      (e: unknown) => e,
    )
    expect(err).toBeInstanceOf(InsertSourceLoadError)
    expect((err as InsertSourceLoadError).kind).toBe('invalid')
  })
})

describe('insertPdfBytes', () => {
  it('round-trips a 3-page source into a 5-page target: 8 pages, order and content kept', async () => {
    const target = await makeLabeledPdf(['T1', 'T2', 'T3', 'T4', 'T5'], [300, 400])
    const source = await makeLabeledPdf(['S1', 'S2', 'S3'], [200, 300])
    // Insert after target page 2 (index 1)
    const { merged, count } = await insertPdfBytes(target, source, 1)
    expect(count).toBe(3)

    const out = await PDFDocument.load(merged)
    expect(out.getPageCount()).toBe(8)
    // Sizes fingerprint the page origin; text proves the content streams moved
    expect(out.getPages().map((p) => p.getWidth())).toEqual([
      300, 300, 200, 200, 200, 300, 300, 300,
    ])
    expect(await pageTexts(merged)).toEqual(['T1', 'T2', 'S1', 'S2', 'S3', 'T3', 'T4', 'T5'])
  })

  it('inserts a picked page subset in the given order at the front', async () => {
    const target = await makeLabeledPdf(['T1'])
    const source = await makeLabeledPdf(['S1', 'S2', 'S3'])
    // Pages 3 and 1 (0-based 2, 0), duplicated and one out of range on top
    const { merged, count } = await insertPdfBytes(target, source, -1, [2, 0, 0, 9])
    expect(count).toBe(2)
    expect(await pageTexts(merged)).toEqual(['S3', 'S1', 'T1'])
  })

  it('carries source annotations, including links that point outside the moved set', async () => {
    const doc = await PDFDocument.create()
    const p1 = doc.addPage([200, 300])
    const p2 = doc.addPage([200, 300])
    const square = doc.context.obj({
      Type: 'Annot',
      Subtype: 'Square',
      Rect: [10, 260, 60, 290],
      C: [1, 0, 0],
    })
    // Internal transition to page 2 — page 2 is NOT part of the inserted set
    const link = doc.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [10, 200, 60, 220],
      Dest: [p2.ref, PDFName.of('Fit')],
    })
    p1.node.set(
      PDFName.of('Annots'),
      doc.context.obj([doc.context.register(square), doc.context.register(link)]),
    )
    const source = await doc.save({ useObjectStreams: false })
    const target = await makeLabeledPdf(['T1'])

    const { merged } = await insertPdfBytes(target, source, -1, [0])
    const out = await PDFDocument.load(merged)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Square', 'Link'])
    // The GoTo destination still resolves to an object in the merged file (the
    // copied page lives outside the page tree — an inert link, not corruption)
    const dest = annots[1]!.lookup(PDFName.of('Dest'), PDFArray)
    expect(dest.get(0)).toBeInstanceOf(PDFRef)
    // The merged bytes parse cleanly in a third-party viewer engine (pdf.js)
    const task = getDocument({ data: merged.slice() })
    try {
      const pdfJsDoc = await task.promise
      expect(pdfJsDoc.numPages).toBe(2)
      const pdfJsAnnots = await (await pdfJsDoc.getPage(1)).getAnnotations()
      expect(pdfJsAnnots.map((a) => a.subtype)).toEqual(['Square', 'Link'])
    } finally {
      await task.destroy()
    }
  })

  it('refuses an encrypted source without touching the target bytes', async () => {
    const target = await makeLabeledPdf(['T1'])
    const err = await insertPdfBytes(target, new Uint8Array(ENCRYPTED_PDF), -1).catch(
      (e: unknown) => e,
    )
    expect(err).toBeInstanceOf(InsertSourceLoadError)
    expect((err as InsertSourceLoadError).kind).toBe('encrypted')
  })

  it('refuses a broken source with the invalid kind', async () => {
    const target = await makeLabeledPdf(['T1'])
    const err = await insertPdfBytes(
      target,
      new Uint8Array(Buffer.from('%PDF-1.4 garbage')),
      -1,
    ).catch((e: unknown) => e)
    expect(err).toBeInstanceOf(InsertSourceLoadError)
    expect((err as InsertSourceLoadError).kind).toBe('invalid')
  })
})

// ── IPC wiring: pick → preview model → in-place insert ─────────────────────

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()
/** Result the mocked native open dialog returns next */
const dialogResult = vi.hoisted(() => ({ value: { canceled: true, filePaths: [] as string[] } }))

let nextWcId = 1
let lastWcId = 0

vi.mock('electron', () => ({
  app: {
    on: vi.fn(),
    whenReady: vi.fn(() => new Promise(() => {})),
    getPath: () => '/tmp/airy-pdf-test-user-data',
  },
  dialog: {
    showOpenDialog: vi.fn(async () => dialogResult.value),
  },
  shell: { showItemInFolder: vi.fn() },
  BrowserWindow: Object.assign(class {}, {
    fromWebContents: () => null,
    getFocusedWindow: () => null,
  }),
  WebContentsView: class {
    webContents = {
      id: (lastWcId = nextWcId++),
      once: vi.fn(),
      on: vi.fn(),
      setWindowOpenHandler: vi.fn(),
    }
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    on: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    removeHandler: vi.fn(),
  },
}))

import { createPdfView } from '../src/main/pdf-main'
import { PDF_CHANNELS } from '../src/shared/ipc'
import type { InsertPdfPickResult, InsertPdfResult } from '../src/shared/ipc'

let dir: string | null = null

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = null
  dialogResult.value = { canceled: true, filePaths: [] }
})

describe('insert-pdf IPC flow', () => {
  it('pick returns the preview model, insert writes the merged file in place', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-insert-ipc-'))
    const targetPath = join(dir, 'target.pdf')
    writeFileSync(targetPath, await makeLabeledPdf(['T1', 'T2', 'T3', 'T4', 'T5'], [300, 400]))
    const sourcePath = join(dir, 'source.pdf')
    writeFileSync(sourcePath, await makeLabeledPdf(['S1', 'S2', 'S3'], [200, 300]))
    createPdfView(targetPath)
    const wc = { sender: { id: lastWcId } }

    dialogResult.value = { canceled: false, filePaths: [sourcePath] }
    const picked = (await handlers.get(PDF_CHANNELS.insertPdfPick)!(wc, {
      path: targetPath,
    })) as InsertPdfPickResult
    expect(picked.ok).toBe(true)
    if (picked.ok && !('canceled' in picked)) {
      expect(picked.name).toBe('source.pdf')
      expect(picked.pages).toEqual([
        { width: 200, height: 300 },
        { width: 200, height: 300 },
        { width: 200, height: 300 },
      ])
    }

    // Insert source pages 1 and 3 after target page 2
    const inserted = (await handlers.get(PDF_CHANNELS.insertPdf)!(wc, {
      path: targetPath,
      afterPageIndex: 1,
      pages: [0, 2],
    })) as InsertPdfResult
    expect(inserted).toEqual({ ok: true, insertedCount: 2 })

    const merged = new Uint8Array(readFileSync(targetPath))
    expect(await pageTexts(merged)).toEqual(['T1', 'T2', 'S1', 'S3', 'T3', 'T4', 'T5'])
  })

  it('refuses an insert when no source was picked, leaving the target untouched', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-insert-nopick-'))
    const targetPath = join(dir, 'target.pdf')
    const bytes = await makeLabeledPdf(['T1'])
    writeFileSync(targetPath, bytes)
    createPdfView(targetPath)

    const result = (await handlers.get(PDF_CHANNELS.insertPdf)!(
      { sender: { id: lastWcId } },
      { path: targetPath, afterPageIndex: -1, pages: [0] },
    )) as InsertPdfResult
    expect(result.ok).toBe(false)
    expect(readFileSync(targetPath)).toEqual(Buffer.from(bytes))
  })

  it('a pick is bound to its view: another view cannot insert it', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-insert-crossview-'))
    const targetPath = join(dir, 'target.pdf')
    const bytes = await makeLabeledPdf(['T1'])
    writeFileSync(targetPath, bytes)
    const sourcePath = join(dir, 'source.pdf')
    writeFileSync(sourcePath, await makeLabeledPdf(['S1']))
    createPdfView(targetPath)
    const picker = { sender: { id: lastWcId } }

    dialogResult.value = { canceled: false, filePaths: [sourcePath] }
    await handlers.get(PDF_CHANNELS.insertPdfPick)!(picker, { path: targetPath })

    // A second view on the same file has no remembered pick
    createPdfView(targetPath)
    const other = { sender: { id: lastWcId } }
    const result = (await handlers.get(PDF_CHANNELS.insertPdf)!(other, {
      path: targetPath,
      afterPageIndex: -1,
    })) as InsertPdfResult
    expect(result.ok).toBe(false)
    expect(readFileSync(targetPath)).toEqual(Buffer.from(bytes))
  })

  it('reports a categorized error for a broken source and keeps the target intact', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-insert-broken-'))
    const targetPath = join(dir, 'target.pdf')
    const bytes = await makeLabeledPdf(['T1'])
    writeFileSync(targetPath, bytes)
    const brokenPath = join(dir, 'broken.pdf')
    writeFileSync(brokenPath, Buffer.from('this is not a pdf'))
    createPdfView(targetPath)
    const wc = { sender: { id: lastWcId } }

    dialogResult.value = { canceled: false, filePaths: [brokenPath] }
    const picked = (await handlers.get(PDF_CHANNELS.insertPdfPick)!(wc, {
      path: targetPath,
    })) as InsertPdfPickResult
    expect(picked.ok).toBe(false)
    if (!picked.ok && 'kind' in picked) expect(picked.kind).toBe('invalid')

    const inserted = (await handlers.get(PDF_CHANNELS.insertPdf)!(wc, {
      path: targetPath,
      afterPageIndex: -1,
    })) as InsertPdfResult
    expect(inserted.ok).toBe(false)
    expect(readFileSync(targetPath)).toEqual(Buffer.from(bytes))
  })

  it('reports the encrypted kind for a password-protected source', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-insert-encrypted-'))
    const targetPath = join(dir, 'target.pdf')
    const bytes = await makeLabeledPdf(['T1'])
    writeFileSync(targetPath, bytes)
    const lockedPath = join(dir, 'locked.pdf')
    writeFileSync(lockedPath, ENCRYPTED_PDF)
    createPdfView(targetPath)

    dialogResult.value = { canceled: false, filePaths: [lockedPath] }
    const picked = (await handlers.get(PDF_CHANNELS.insertPdfPick)!(
      { sender: { id: lastWcId } },
      { path: targetPath },
    )) as InsertPdfPickResult
    expect(picked.ok).toBe(false)
    if (!picked.ok && 'kind' in picked) expect(picked.kind).toBe('encrypted')
    expect(readFileSync(targetPath)).toEqual(Buffer.from(bytes))
  })

  it('a canceled pick dialog is a cancel, not a failure', async () => {
    dir = mkdtempSync(join(tmpdir(), 'pdf-insert-cancel-'))
    const targetPath = join(dir, 'target.pdf')
    writeFileSync(targetPath, await makeLabeledPdf(['T1']))
    createPdfView(targetPath)

    dialogResult.value = { canceled: true, filePaths: [] }
    const picked = (await handlers.get(PDF_CHANNELS.insertPdfPick)!(
      { sender: { id: lastWcId } },
      { path: targetPath },
    )) as InsertPdfPickResult
    expect(picked).toEqual({ ok: true, canceled: true })
  })
})

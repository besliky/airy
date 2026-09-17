/// PDF export: renders the print HTML (laid out by the renderer) in a hidden
/// scripting-disabled window and writes webContents.printToPDF's output where
/// the save dialog points.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { BrowserWindow, dialog } from 'electron'

import { showSaveDialogWithMemory } from '@airy-office/electron-utils'

import { evenPageRanges, stitchPlan, type PageVariant } from './pdf-page-variants'

import type { IpcMainInvokeEvent, WebContents } from 'electron'
import type { PDFDocument } from 'pdf-lib'
import type {
  WorkbookExportPdfRequest,
  WorkbookExportPdfResult,
  WorkbookPrintPreviewResult,
  WorkbookPrintResult,
} from '../shared/desktop-api'

export async function exportPdf(
  event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookExportPdfResult> {
  const parent = BrowserWindow.fromWebContents(event.sender)
  const dialogOptions = {
    defaultPath: request.fileName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  }
  const selection = await showSaveDialogWithMemory(dialog, parent, dialogOptions)
  if (selection.canceled || !selection.filePath) return { canceled: true }

  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-pdf-'))
  const htmlPath = join(workDir, 'print.html')
  const window = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    const pdf = await renderPdf(window.webContents, request)
    await writeFile(selection.filePath, pdf)
    return { canceled: false, path: selection.filePath }
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

interface TemplatePair {
  readonly headerTemplate?: string | undefined
  readonly footerTemplate?: string | undefined
}

/// One printToPDF pass; `pageRanges` undefined prints every page. Chromium
/// falls back to its own date/title header when a template is missing, so
/// both templates are always passed once headers/footers are shown.
async function printPass(
  contents: WebContents,
  request: WorkbookExportPdfRequest,
  templates: TemplatePair | undefined,
  pageRanges?: string,
): Promise<Buffer> {
  return contents.printToPDF({
    landscape: request.landscape,
    pageSize: request.pageSize,
    margins: request.margins,
    scale: request.scale,
    printBackground: true,
    ...(pageRanges === undefined ? {} : { pageRanges }),
    ...(templates
      ? {
          displayHeaderFooter: true,
          headerTemplate: templates.headerTemplate ?? '<span></span>',
          footerTemplate: templates.footerTemplate ?? '<span></span>',
        }
      : {}),
  })
}

/// Header/footer templates of the base (odd) pass: present once the request
/// carries any header/footer, so Chromium never substitutes its own.
function oddTemplatesFor(request: WorkbookExportPdfRequest): TemplatePair | undefined {
  return request.headerTemplate !== undefined || request.footerTemplate !== undefined
    ? { headerTemplate: request.headerTemplate, footerTemplate: request.footerTemplate }
    : undefined
}

/// Chromium prints one header/footer template pair for every page. Excel's
/// differentFirst / differentOddEven need extra passes — page 1 with the
/// first-page templates, the even pages with the even ones — stitched into
/// the odd-page print by page index (pdf-lib). Without variants the odd pass
/// is the whole export (single-pass fast path).
async function renderPdf(
  contents: WebContents,
  request: WorkbookExportPdfRequest,
): Promise<Buffer> {
  const oddTemplates = oddTemplatesFor(request)
  const flags = {
    hasFirst: request.firstPage !== undefined,
    hasEven: request.evenPages !== undefined,
  }
  const showHeaderFooter = oddTemplates !== undefined || flags.hasFirst || flags.hasEven
  const odd = await printPass(
    contents,
    request,
    showHeaderFooter ? (oddTemplates ?? {}) : undefined,
  )
  if (!flags.hasFirst && !flags.hasEven) return odd

  const { PDFDocument: PdfDocument } = await import('pdf-lib')
  const oddDocument = await PdfDocument.load(odd)
  const total = oddDocument.getPageCount()
  const passes: Partial<Record<PageVariant, PDFDocument>> = { odd: oddDocument }
  if (request.firstPage !== undefined && total >= 1) {
    const first = await printPass(contents, request, request.firstPage, '1')
    passes.first = await PdfDocument.load(first)
  }
  const evenRanges = evenPageRanges(total)
  if (request.evenPages !== undefined && evenRanges !== '') {
    const even = await printPass(contents, request, request.evenPages, evenRanges)
    passes.even = await PdfDocument.load(even)
  }
  const merged = await PdfDocument.create()
  for (const step of stitchPlan(total, flags)) {
    // A pass that came back with fewer pages than planned (Chromium and
    // pdf-lib disagreeing about a range) falls back to the odd print of
    // that page rather than failing the export.
    const source = passes[step.source]
    const [page] =
      source !== undefined && step.index < source.getPageCount()
        ? await merged.copyPages(source, [step.index])
        : await merged.copyPages(oddDocument, [step.page - 1])
    if (page) merged.addPage(page)
  }
  return Buffer.from(await merged.save())
}

/// Shared scaffolding for the print dialog's channels: the same hidden
/// scripting-disabled window the export uses, handed to a callback instead
/// of a save dialog + file write.
async function withPrintWindow<T>(
  request: WorkbookExportPdfRequest,
  run: (contents: WebContents) => Promise<T>,
): Promise<T> {
  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-print-'))
  const htmlPath = join(workDir, 'print.html')
  const window = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    return await run(window.webContents)
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

/// Page count from raw PDF bytes. Chromium's printToPDF leaves the page
/// object dictionaries uncompressed (only content streams are deflated), so
/// scanning for `/Type /Page` — excluding the `/Type /Pages` tree nodes —
/// avoids a full pdf-lib parse of the preview on every option change.
export function countPdfPages(bytes: Uint8Array): number {
  const text = Buffer.from(bytes).toString('latin1')
  let count = 0
  for (const _match of text.matchAll(/\/Type\s*\/Page(?![A-Za-z])/g)) count += 1
  return count
}

/// Page count with a belt-and-suspenders fallback: a producer that compresses
/// the page dictionaries yields no scan hits, so fall back to a real parse.
async function pageCountOf(pdf: Buffer): Promise<number> {
  const scanned = countPdfPages(pdf)
  if (scanned > 0) return scanned
  const { PDFDocument: PdfDocument } = await import('pdf-lib')
  return (await PdfDocument.load(pdf)).getPageCount()
}

/// Print dialog preview: the request's page count. Nothing touches disk, no
/// save dialog appears, and no PDF bytes cross the IPC boundary — the dialog
/// previews the print HTML itself in its iframe; only the count needs the
/// main-side printToPDF pass. The variant stitching changes which template
/// each page carries, not how many pages print, so the base (odd) pass alone
/// carries the count.
export async function previewPrint(
  _event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookPrintPreviewResult> {
  try {
    return await withPrintWindow(request, async (contents) => {
      const oddTemplates = oddTemplatesFor(request)
      const showHeaderFooter =
        oddTemplates !== undefined ||
        request.firstPage !== undefined ||
        request.evenPages !== undefined
      const pdf = await printPass(
        contents,
        request,
        showHeaderFooter ? (oddTemplates ?? {}) : undefined,
      )
      return { ok: true, pageCount: await pageCountOf(pdf) }
    })
  } catch (err) {
    return { ok: false, error: String(err) }
  }
}

/// Print dialog Print: the system print dialog over the print HTML, with
/// the request's paper size, orientation, margins, and scale. Header and
/// footer templates only render through printToPDF (Chromium's limitation),
/// so they stay a PDF-export feature. Resolves when the system dialog is
/// dismissed; ok=false without an error means the user canceled there.
export async function printWorkbook(
  _event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookPrintResult> {
  try {
    return await withPrintWindow(request, async (contents) => {
      const printed = await new Promise<{ ok: boolean; error?: string }>((resolve) => {
        contents.print(
          {
            printBackground: true,
            landscape: request.landscape,
            pageSize: request.pageSize,
            margins: {
              marginType: 'custom',
              top: request.margins.top,
              bottom: request.margins.bottom,
              left: request.margins.left,
              right: request.margins.right,
            },
            scaleFactor: Math.round(request.scale * 100),
          },
          (success, failureReason) => {
            resolve({
              ok: success,
              ...(failureReason && !/cancel/i.test(failureReason) ? { error: failureReason } : {}),
            })
          },
        )
      })
      return printed satisfies WorkbookPrintResult
    })
  } catch (err) {
    return { ok: false, error: String(err) }
  }
}

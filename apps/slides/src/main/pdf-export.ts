import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildPrintDocumentHtml } from '../shared/print-html'

export interface PdfExportWindow {
  loadFile(path: string): Promise<void>
  webContents: {
    printToPDF(options: Electron.PrintToPDFOptions): Promise<Buffer>
  }
  destroy(): void
}

export interface ExportSlidesPdfOptions {
  /** one entry per exported slide: inline vector SVG (preferred) or raster PNG fallback */
  pages: Array<{ svg?: string; pngBase64?: string }>
  widthPx: number
  heightPx: number
  filePath: string
  /**
   * Page layout: 'full' paints one slide per PDF page (default); the other
   * layouts reuse the print sheet's exact page assembly (A4 portrait).
   */
  layout?: 'full' | 'notes' | 'handout2' | 'handout3'
  /** Per-slide speaker notes for the 'notes' layout (same order as pages) */
  notes?: string[]
  createWindow(): PdfExportWindow
  openExportedPdf(path: string): void
}

export interface ExportSlidesPdfResult {
  ok: boolean
  path?: string
  error?: string
}

/** Inline page body: the vector slide fills the page; raster pages embed the bitmap. */
export function pdfPageBody(page: { svg?: string; pngBase64?: string }): string {
  if (page.svg) return page.svg
  return `<img src="data:image/png;base64,${page.pngBase64 ?? ''}" alt="">`
}

export function buildPdfExportHtml(
  pages: Array<{ svg?: string; pngBase64?: string }>,
  widthIn: number,
  heightIn: number,
): string {
  return `<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: ${widthIn}in ${heightIn}in; margin: 0; }
html, body { margin: 0; padding: 0; }
.page { width: ${widthIn}in; height: ${heightIn}in; overflow: hidden; page-break-after: always; }
.page:last-child { page-break-after: auto; }
.page img { display: block; width: 100%; height: 100%; }
.page svg { display: block; width: 100%; height: 100%; }
</style></head><body>${pages
    .map((p) => `<div class="page">${pdfPageBody(p)}</div>`)
    .join('')}</body></html>`
}

/** Export the deck via an app-owned temporary HTML file (printToPDF: SVG text stays selectable). */
export async function exportSlidesPdf({
  pages,
  widthPx,
  heightPx,
  filePath,
  layout = 'full',
  notes,
  createWindow,
  openExportedPdf,
}: ExportSlidesPdfOptions): Promise<ExportSlidesPdfResult> {
  // PDF page size: 'full' keeps the slide ratio at the 7.5in print height; the
  // notes/handout layouts print on A4 portrait exactly like the print sheet
  const heightIn = layout === 'full' ? 7.5 : 11.69
  const widthIn =
    layout === 'full' ? Math.round((widthPx / heightPx) * heightIn * 1000) / 1000 : 8.27
  const html =
    layout === 'full'
      ? buildPdfExportHtml(pages, widthIn, heightIn)
      : // same assembly the print preview/print job use, with the vector slides
        // inlined where the preview puts its bitmap thumbnails
        buildPrintDocumentHtml({
          srcs: pages.map((p) => (p.svg ? '' : `data:image/png;base64,${p.pngBase64 ?? ''}`)),
          svgs: pages.map((p) => p.svg),
          ratio: widthPx / heightPx,
          layout,
          ...(layout === 'notes' ? { notes: notes ?? [] } : {}),
        })
  const win = createWindow()
  let tempDir: string | null = null
  try {
    tempDir = await mkdtemp(join(tmpdir(), 'airy-slides-pdf-'))
    const htmlPath = join(tempDir, 'slides.html')
    await writeFile(htmlPath, html, 'utf8')
    await win.loadFile(htmlPath)
    // The window is scripting-disabled (javascript: false, like sheets'
    // pdf-export), so no fonts/images-ready probe runs: loadFile resolves at
    // onload (all data:-URL images loaded) and printToPDF rasterizes the
    // decoded result — the sheets export path proves this renders
    // SVG text and bitmaps correctly.
    const pdf = await win.webContents.printToPDF({
      landscape: false, // The page size is already landscape (width > height); passing landscape would rotate a second time
      printBackground: true,
      pageSize: { width: widthIn, height: heightIn },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      preferCSSPageSize: false,
    })
    await writeFile(filePath, pdf)
    openExportedPdf(filePath)
    return { ok: true, path: filePath }
  } catch (err) {
    return { ok: false, error: String(err) }
  } finally {
    try {
      win.destroy()
    } finally {
      if (tempDir) await rm(tempDir, { recursive: true, force: true })
    }
  }
}

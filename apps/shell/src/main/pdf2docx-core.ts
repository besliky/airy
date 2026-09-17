/**
 * In-process PDF → docx conversion core, shared by the utilityProcess worker
 * entry (pdf2docx-worker.ts, the normal Electron path) and the non-Electron
 * fallback in pdf2docx-local.ts (tests, plain node). Everything here must stay
 * free of Electron imports so both hosts can run it.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { convertPdfToDocx } from '../../../../packages/pdf2docx/src'
import type { ConvertResult, OcrEngine, PdfiumModule } from '../../../../packages/pdf2docx/src'
import {
  createVisionOcrEngine,
  createWindowsOcrEngine,
} from '../../../../packages/pdf2docx/src/ocr-vision'
import { pdfiumWasmPath } from '../../../pdf/src/main/wasm-path'

/**
 * Local OCR engine for scanned pages (platform system OCR; see
 * packages/pdf2docx/src/ocr.ts) — macOS Vision on darwin, Windows.Media.Ocr
 * on win32. Optional by design: when the helper binary is absent (Linux, or
 * a build without it) the engine resolves null and scanned pages keep the
 * full-page-image fallback.
 *
 * Packaged: Resources/ocr/<helper> (electron-builder extraResources).
 * Dev: the compiled helper in the repo (packages/pdf2docx/ocr-helper/).
 */
let ocrEngine: OcrEngine | null | undefined
function ensureOcrEngine(): OcrEngine | null {
  if (ocrEngine !== undefined) return ocrEngine
  const here = dirname(fileURLToPath(import.meta.url))
  const helper = process.platform === 'darwin' ? 'vision-ocr' : 'win-ocr.exe'
  const create = process.platform === 'darwin' ? createVisionOcrEngine : createWindowsOcrEngine
  const candidates = [
    ...(process.resourcesPath ? [join(process.resourcesPath, 'ocr', helper)] : []),
    join(here, '../../../../packages/pdf2docx/ocr-helper', helper),
  ]
  ocrEngine = null
  for (const path of candidates) {
    const engine = create(path)
    if (engine) {
      ocrEngine = engine
      break
    }
  }
  return ocrEngine
}

let pdfiumPromise: Promise<PdfiumModule> | null = null

/** Load the wasm bytes ourselves: neither host may rely on the package's own
 *  file resolution (see apps/pdf text-edit.ts). Shared by both the worker and
 *  the in-process fallback so the wasm singleton lives per process. */
export function ensurePdfium(): Promise<PdfiumModule> {
  pdfiumPromise ??= (async () => {
    const { init } = (await import('@embedpdf/pdfium')) as unknown as {
      init(overrides: object): Promise<object>
    }
    const raw = readFileSync(pdfiumWasmPath())
    // exact slice: Buffer.buffer may be a shared pool larger than the file
    const wasmBinary = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength)
    // thisProgram: emscripten's synthetic environ writes process.argv[1] via
    // ASCII-asserting stringToAscii; a document path with CJK characters handed
    // to the packaged app by a Windows file association aborts init (same fix
    // as apps/pdf/src/main/text-edit.ts loadPdfium)
    const wrapped = (await init({ wasmBinary, thisProgram: 'airy-pdf' })) as {
      pdfium?: unknown
    }
    const m = (wrapped.pdfium ?? wrapped) as PdfiumModule & { _PDFiumExt_Init(): void }
    m._PDFiumExt_Init()
    return m
  })()
  return pdfiumPromise
}

/** Convert a PDF file on disk to docx bytes, fully locally, in this process. */
export async function convertPdfFileToDocxInProcess(
  pdfPath: string,
  onProgress?: (page: number, total: number) => void,
  password?: string,
): Promise<ConvertResult> {
  const pdfium = await ensurePdfium()
  const bytes = readFileSync(pdfPath)
  const pdf = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const ocr = ensureOcrEngine()
  return convertPdfToDocx(pdf, {
    pdfium,
    ...(ocr ? { ocr } : {}),
    ...(onProgress !== undefined ? { onProgress } : {}),
    ...(password !== undefined ? { password } : {}),
  })
}

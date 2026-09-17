/**
 * utilityProcess entry for local PDF → docx conversion. The wasm conversion is
 * CPU-bound and used to stall the shell main between page callbacks (menus,
 * IPC); running it here keeps the main process responsive. Bundled next to the
 * main bundle (electron.vite.config.ts extra main entry). The protocol is a
 * small promise RPC owned by pdf2docx-local.ts — see convertPdfFileToDocxLocal.
 */
import { PdfLoadError } from '../../../../packages/pdf2docx/src'
import { convertPdfFileToDocxInProcess } from './pdf2docx-core'

interface ConvertRequest {
  id: number
  type: 'convert'
  pdfPath: string
  password?: string
}

if (process.parentPort) {
  process.parentPort.on('message', (event) => {
    const request = event.data as ConvertRequest
    if (request?.type !== 'convert') return
    void (async () => {
      try {
        const result = await convertPdfFileToDocxInProcess(
          request.pdfPath,
          // progress is per-conversion; the client correlates by id
          (page, total) =>
            process.parentPort!.postMessage({ id: request.id, type: 'progress', page, total }),
          request.password,
        )
        process.parentPort!.postMessage({ id: request.id, type: 'done', result })
      } catch (error) {
        // Errors crossing the process boundary lose their class identity;
        // carry enough to rebuild PdfLoadError on the main side (its code
        // drives the password-retry flow and the user-facing dialog).
        const serialized =
          error instanceof PdfLoadError
            ? {
                message: error.message,
                pdfLoadError: { code: error.code, pdfiumError: error.pdfiumError },
              }
            : { message: error instanceof Error ? error.message : String(error) }
        process.parentPort!.postMessage({ id: request.id, type: 'error', ...serialized })
      }
    })()
  })
}

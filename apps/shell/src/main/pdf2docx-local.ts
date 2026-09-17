/**
 * Local PDF → Word conversion for the shell's pdf tabs (pdf2docx P4).
 * The CPU-bound wasm pipeline runs in an Electron utilityProcess
 * (pdf2docx-worker.ts) so page-by-page conversion no longer stalls the main
 * process between callbacks (menus, IPC, window events). Without Electron
 * (vitest, plain node) it degrades to the shared in-process core. The pdf →
 * pptx/xlsx exporters still convert in the main process and share this
 * module's ensurePdfium singleton. Imported by relative path (like the other
 * sibling app modules) so the bundled shell main carries the package inline.
 */
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { PdfLoadError } from '../../../../packages/pdf2docx/src'
import type { ConvertResult, PdfLoadErrorCode } from '../../../../packages/pdf2docx/src'
import { convertPdfFileToDocxInProcess, ensurePdfium } from './pdf2docx-core'

export type { ConvertResult, PageResult } from '../../../../packages/pdf2docx/src'
export { PdfLoadError } from '../../../../packages/pdf2docx/src'
export { ensurePdfium }

/** Minimal shape of Electron's utilityProcess child used by the client. */
export interface PdfWorkerChild {
  postMessage(message: unknown): void
  on(event: 'message', listener: (message: unknown) => void): PdfWorkerChild
  once(event: 'exit', listener: (code: number) => void): PdfWorkerChild
  kill(): void
}
export type PdfWorkerFork = (modulePath: string) => PdfWorkerChild

interface WorkerMessage {
  id: number
  type: 'progress' | 'done' | 'error'
  page?: number
  total?: number
  result?: ConvertResult
  message?: string
  pdfLoadError?: { code: PdfLoadErrorCode; pdfiumError: number }
}

/** Live workers, so a crash or app quit cannot leave stragglers behind. */
const liveWorkers = new Set<PdfWorkerChild>()
let conversionSequence = 0

/** The bundled worker entry sits next to the main bundle (electron.vite extra
 *  main entry); resolved relative to this file so dev and packaged both work. */
function workerModulePath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), 'pdf2docx-worker.js')
}

/**
 * Run one conversion in a freshly forked worker via the promise RPC. A fresh
 * process per conversion keeps failure isolation (a crashed worker rejects
 * only its own conversion) and makes the extra wasm init irrelevant next to
 * the conversion itself. Exported for tests, which fake `fork` at this
 * boundary.
 */
export function convertPdfFileToDocxVia(
  fork: PdfWorkerFork,
  pdfPath: string,
  onProgress?: (page: number, total: number) => void,
  password?: string,
): Promise<ConvertResult> {
  const id = ++conversionSequence
  const child = fork(workerModulePath())
  liveWorkers.add(child)
  return new Promise<ConvertResult>((resolve, reject) => {
    let settled = false
    const settle = (run: () => void) => {
      if (settled) return
      settled = true
      liveWorkers.delete(child)
      run()
    }
    child.on('message', (raw) => {
      const message = raw as WorkerMessage
      if (message?.id !== id) return
      if (message.type === 'progress') {
        onProgress?.(message.page!, message.total!)
      } else if (message.type === 'done') {
        settle(() => resolve(message.result!))
      } else if (message.type === 'error') {
        settle(() => {
          const detail = message.pdfLoadError
          reject(
            detail
              ? new PdfLoadError(detail.code, detail.pdfiumError)
              : new Error(message.message ?? 'PDF conversion failed'),
          )
        })
      }
    })
    child.once('exit', (code) => {
      settle(() =>
        reject(
          new Error(
            `The PDF conversion worker exited unexpectedly (code ${code}). ` +
              'Retry the conversion; if it persists, the PDFium runtime may be missing or damaged.',
          ),
        ),
      )
    })
    child.postMessage({
      id,
      type: 'convert',
      pdfPath,
      ...(password !== undefined ? { password } : {}),
    })
  })
}

/** Electron's utilityProcess.fork, when actually running under Electron. */
async function electronFork(): Promise<PdfWorkerFork | null> {
  try {
    const electron = (await import('electron')) as unknown as
      typeof import('electron') | { default?: unknown }
    const candidate = (electron as { utilityProcess?: { fork?: unknown } }).utilityProcess
    if (candidate && typeof candidate.fork === 'function') {
      const fork = candidate.fork.bind(candidate) as (
        modulePath: string,
        args: readonly string[],
        options: { serviceName: string },
      ) => PdfWorkerChild
      return (modulePath: string) => fork(modulePath, [], { serviceName: 'pdf2docx' })
    }
    return null
  } catch {
    return null
  }
}

/** Kill any in-flight conversion workers (app quit, teardown). */
export function disposePdfConversionWorkers(): void {
  for (const child of liveWorkers) {
    try {
      child.kill()
    } catch {
      /* already gone */
    }
  }
  liveWorkers.clear()
}

/**
 * Convert a PDF file on disk to docx bytes, fully locally. Runs the wasm
 * pipeline in a utilityProcess under Electron; without Electron (tests, plain
 * node) it runs the identical shared core in-process.
 */
export async function convertPdfFileToDocxLocal(
  pdfPath: string,
  onProgress?: (page: number, total: number) => void,
  password?: string,
): Promise<ConvertResult> {
  const fork = await electronFork()
  if (!fork) return convertPdfFileToDocxInProcess(pdfPath, onProgress, password)
  return convertPdfFileToDocxVia(fork, pdfPath, onProgress, password)
}

/**
 * Password retry loop around a conversion attempt (P23). Runs `convert`
 * without a password first; on PdfLoadError('password-required') asks
 * `promptPassword` (retry=true once a submitted password was rejected) and
 * re-runs until it succeeds, a different error is thrown, or the prompt
 * returns null (user cancelled) → resolves null. Pure state machine, exported
 * separately from the UI so tests can drive it with fakes or the real
 * converter.
 */
export async function convertWithPasswordRetry<T>(
  convert: (password: string | undefined) => Promise<T>,
  promptPassword: (retry: boolean) => Promise<string | null>,
): Promise<T | null> {
  let password: string | undefined
  for (;;) {
    try {
      return await convert(password)
    } catch (err) {
      if (!(err instanceof PdfLoadError) || err.code !== 'password-required') throw err
      const entered = await promptPassword(password !== undefined)
      if (entered === null) return null
      password = entered
    }
  }
}

/**
 * Local conversion with an interactive password prompt: like
 * convertPdfFileToDocxLocal but encrypted PDFs ask the user for the password
 * (looping on wrong entries) instead of failing. Resolves null when the user
 * cancels the prompt.
 */
export function convertPdfFileToDocxLocalWithPrompt(
  pdfPath: string,
  promptPassword: (retry: boolean) => Promise<string | null>,
  onProgress?: (page: number, total: number) => void,
): Promise<ConvertResult | null> {
  return convertWithPasswordRetry(
    (password) => convertPdfFileToDocxLocal(pdfPath, onProgress, password),
    promptPassword,
  )
}

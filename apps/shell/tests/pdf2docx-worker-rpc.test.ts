import { describe, expect, it, vi } from 'vitest'
import { EventEmitter } from 'node:events'
import {
  convertPdfFileToDocxVia,
  disposePdfConversionWorkers,
  PdfLoadError,
  type PdfWorkerChild,
} from '../src/main/pdf2docx-local'

/**
 * Promise-RPC client over the pdf2docx utilityProcess worker, tested with a
 * fake fork at the boundary: progress forwarding, result delivery, PdfLoadError
 * identity reconstruction (errors lose their class across the process
 * boundary), and worker-crash rejection. The worker side itself is the shared
 * in-process core covered by pdf2docx-local-argv.test.ts.
 */
function fakeWorker() {
  const emitter = new EventEmitter() as EventEmitter & { sent: unknown[] }
  emitter.sent = []
  const child: PdfWorkerChild = {
    postMessage: (message: unknown) => void emitter.sent.push(message),
    on: (event, listener) => {
      emitter.on(event, listener)
      return child
    },
    once: (event, listener) => {
      emitter.once(event, listener)
      return child
    },
    kill: () => emitter.emit('exit', 0),
  }
  return {
    child,
    sent: emitter.sent,
    emit: (message: unknown) => emitter.emit('message', message),
    crash: () => emitter.emit('exit', 7),
  }
}

describe('convertPdfFileToDocxVia (worker RPC client)', () => {
  it('sends the convert request and forwards progress, then resolves the result', async () => {
    const worker = fakeWorker()
    const fork = vi.fn(() => worker.child)
    const onProgress = vi.fn()
    const promise = convertPdfFileToDocxVia(fork, '/docs/a.pdf', onProgress, 'pw')
    expect(fork).toHaveBeenCalledWith(expect.stringContaining('pdf2docx-worker.js'))
    expect(worker.sent).toEqual([
      { id: 1, type: 'convert', pdfPath: '/docs/a.pdf', password: 'pw' },
    ])

    worker.emit({ id: 1, type: 'progress', page: 2, total: 5 })
    worker.emit({ id: 1, type: 'progress', page: 5, total: 5 })
    worker.emit({ id: 1, type: 'done', result: { docx: new Uint8Array([1]), pages: 5 } })

    await expect(promise).resolves.toMatchObject({ pages: 5 })
    expect(onProgress).toHaveBeenNthCalledWith(1, 2, 5)
    expect(onProgress).toHaveBeenNthCalledWith(2, 5, 5)
  })

  it('ignores messages for other conversion ids', async () => {
    const worker = fakeWorker()
    const promise = convertPdfFileToDocxVia(() => worker.child, '/docs/a.pdf')
    worker.emit({ id: 99, type: 'done', result: { docx: new Uint8Array(), pages: 1 } })
    worker.emit({ id: 2, type: 'done', result: { docx: new Uint8Array([7]), pages: 3 } })
    await expect(promise).resolves.toMatchObject({ pages: 3 })
  })

  it('rebuilds PdfLoadError with its code so the password-retry flow still matches', async () => {
    const worker = fakeWorker()
    const promise = convertPdfFileToDocxVia(() => worker.child, '/docs/a.pdf')
    worker.emit({
      id: 3,
      type: 'error',
      message: 'PDFium could not load the document (password-required, FPDF error 4)',
      pdfLoadError: { code: 'password-required', pdfiumError: 4 },
    })
    const error = await promise.catch((err: unknown) => err)
    expect(error).toBeInstanceOf(PdfLoadError)
    expect((error as PdfLoadError).code).toBe('password-required')
    expect((error as PdfLoadError).pdfiumError).toBe(4)
  })

  it('rejects with the plain message when the worker reports a generic error', async () => {
    const worker = fakeWorker()
    const promise = convertPdfFileToDocxVia(() => worker.child, '/docs/a.pdf')
    worker.emit({ id: 4, type: 'error', message: 'wasm init failed' })
    await expect(promise).rejects.toThrow('wasm init failed')
  })

  it('rejects with an actionable error when the worker crashes mid-conversion', async () => {
    const worker = fakeWorker()
    const promise = convertPdfFileToDocxVia(() => worker.child, '/docs/a.pdf')
    worker.crash()
    await expect(promise).rejects.toThrow('worker exited unexpectedly (code 7)')
  })

  it('disposePdfConversionWorkers kills live workers without rejecting twice', async () => {
    const worker = fakeWorker()
    const promise = convertPdfFileToDocxVia(() => worker.child, '/docs/a.pdf')
    disposePdfConversionWorkers()
    await expect(promise).rejects.toThrow('worker exited unexpectedly')
  })
})

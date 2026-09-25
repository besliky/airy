import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { PDFDocument } from 'pdf-lib'

/**
 * pdf:save staleness fence (BUG-1730, docs #151 analog): an in-place save must
 * refuse to silently overwrite a file another window/program rewrote after this
 * view last read or saved it. Autosave is refused without any dialog; a manual
 * save raises the Save As / Overwrite / Cancel prompt; the fence's Save As lands
 * the edits on a user-picked copy and never writes the contested original.
 */

interface FakeSender {
  id: number
  isDestroyed: () => boolean
}

type IpcHandler = (event: { sender: FakeSender }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()
const showMessageBox = vi.fn()
const showSaveDialog = vi.fn()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
  on: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  loadFile: ReturnType<typeof vi.fn>
}

let nextWcId = 1
let lastWebContents: FakeWebContents

function makeFakeWebContents(): FakeWebContents {
  const wc: FakeWebContents = {
    id: nextWcId++,
    once: vi.fn(),
    on: vi.fn(),
    setWindowOpenHandler: vi.fn(),
    loadURL: vi.fn(),
    loadFile: vi.fn(),
  }
  lastWebContents = wc
  return wc
}

vi.mock('electron', () => ({
  app: {
    on: vi.fn(),
    whenReady: vi.fn(() => new Promise(() => {})),
    getPath: () => '/tmp/airy-pdf-test-user-data',
  },
  // deferred wrappers: the factory runs before the mock fns below initialize
  dialog: {
    showMessageBox: (...args: unknown[]) => showMessageBox(...args),
    showSaveDialog: (...args: unknown[]) => showSaveDialog(...args),
  },
  shell: {},
  // statics used by the fence: no focused window in tests → headless dialog fallback
  BrowserWindow: Object.assign(vi.fn(), {
    fromWebContents: () => null,
    getFocusedWindow: () => null,
  }),
  WebContentsView: class {
    webContents = makeFakeWebContents()
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

import { setUiLang } from '@airy-office/i18n'
import { PDF_CHANNELS } from '../src/shared/ipc'
import type { SavePdfRequest, SavePdfResult } from '../src/shared/ipc'
import { createPdfView } from '../src/main/pdf-main'

// the fence dialog strings are asserted in English (docs #151 wording)
setUiLang('en')

const tempDirs = new Set<string>()

async function makePdfFile(name: string, marker = 'A'): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), `pdf-save-fence-${marker}-`))
  tempDirs.add(dir)
  const path = join(dir, name)
  writeFileSync(path, await pdfBytes(100))
  return path
}

/** A valid PDF of the given page width — differing widths keep contents distinct */
async function pdfBytes(pageWidth: number): Promise<Buffer> {
  const doc = await PDFDocument.create()
  doc.addPage([pageWidth, 100])
  return Buffer.from(await doc.save({ useObjectStreams: false }))
}

const senderOf = (wcId: number) => ({ id: wcId, isDestroyed: () => false })

const readGranted = (wcId: number, path: string) =>
  handlers.get(PDF_CHANNELS.readFile)?.({ sender: senderOf(wcId) }, path) as Promise<ArrayBuffer>

const saveGranted = (wcId: number, request: SavePdfRequest) =>
  handlers.get(PDF_CHANNELS.save)?.({ sender: senderOf(wcId) }, request) as Promise<SavePdfResult>

const request = (path: string, over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path,
  markups: [
    // a highlight so every save rewrites the file observably
    {
      pageIndex: 0,
      type: 'highlight',
      color: [1, 0, 0],
      quads: [[10, 10, 60, 10, 10, 30, 60, 30]],
    },
  ],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

afterEach(() => {
  vi.clearAllMocks()
  for (const dir of tempDirs) rmSync(dir, { recursive: true, force: true })
  tempDirs.clear()
})

/** open simulation: the renderer always reads the file through main before any save */
const openIn = async (wcId: number, path: string) => {
  const data = await readGranted(wcId, path)
  expect(data.byteLength).toBeGreaterThan(0)
}

describe('pdf:save staleness fence', () => {
  it('saves in place when the file is untouched, and again right after (own write re-stamped)', async () => {
    const path = await makePdfFile('doc.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)
    const before = readFileSync(path)

    const first = await saveGranted(wcId, request(path))
    expect(first).toMatchObject({ ok: true })
    expect(showMessageBox).not.toHaveBeenCalled()
    expect(readFileSync(path)).not.toEqual(before)

    // the save itself refreshed the fence baseline: no false conflict on the next save
    const second = await saveGranted(wcId, request(path))
    expect(second).toMatchObject({ ok: true })
    expect(showMessageBox).not.toHaveBeenCalled()
  })

  it('manual save on an externally modified file raises the dialog; Overwrite lands the save', async () => {
    const path = await makePdfFile('doc.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)

    const external = await pdfBytes(200)
    writeFileSync(path, external)

    showMessageBox.mockResolvedValue({ response: 1 }) // Overwrite
    const result = await saveGranted(wcId, request(path))
    expect(result).toMatchObject({ ok: true })
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    const asked = showMessageBox.mock.calls[0]![0] as { message: string }
    expect(asked.message).toBe('The file has been modified by another program.')
    expect(readFileSync(path)).not.toEqual(external)
  })

  it('dialog Cancel refuses the save and keeps the external file intact', async () => {
    const path = await makePdfFile('doc.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)

    const external = await pdfBytes(200)
    writeFileSync(path, external)

    showMessageBox.mockResolvedValue({ response: 2 }) // Cancel
    const result = await saveGranted(wcId, request(path))
    expect(result).toEqual({ ok: false, reason: 'external-modified' })
    expect(readFileSync(path)).toEqual(external)
  })

  it('autosave on a stale file is refused silently, without any dialog', async () => {
    const path = await makePdfFile('doc.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)

    const external = await pdfBytes(200)
    writeFileSync(path, external)

    const result = await saveGranted(wcId, request(path, { auto: true }))
    expect(result).toEqual({ ok: false, reason: 'external-modified' })
    expect(showMessageBox).not.toHaveBeenCalled()
    expect(showSaveDialog).not.toHaveBeenCalled()
    expect(readFileSync(path)).toEqual(external)
  })

  it('fence Save As lands the edits on the picked copy and never writes the original', async () => {
    const path = await makePdfFile('doc.pdf')
    const copyPath = join(path, '..', 'doc-copy.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)

    const external = await pdfBytes(200)
    writeFileSync(path, external)

    showMessageBox.mockResolvedValue({ response: 0 }) // Save As (default)
    showSaveDialog.mockResolvedValue({ canceled: false, filePath: copyPath })
    const result = await saveGranted(wcId, request(path))
    expect(result).toMatchObject({ ok: true, savedAsPath: copyPath })
    expect(showSaveDialog).toHaveBeenCalledTimes(1)
    // the copy carries the saved document, the contested original is untouched
    expect(existsSync(copyPath)).toBe(true)
    expect(readFileSync(copyPath)).not.toEqual(external)
    expect(readFileSync(path)).toEqual(external)
  })

  it('fence Save As with a canceled picker refuses without writing anything', async () => {
    const path = await makePdfFile('doc.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)

    const external = await pdfBytes(200)
    writeFileSync(path, external)

    showMessageBox.mockResolvedValue({ response: 0 }) // Save As
    showSaveDialog.mockResolvedValue({ canceled: true, filePath: undefined })
    const result = await saveGranted(wcId, request(path))
    expect(result).toEqual({ ok: false, reason: 'external-modified' })
    expect(readFileSync(path)).toEqual(external)
  })

  it('two windows on one file: the stale second window hits the fence the first saved cleanly', async () => {
    // the BUG-1730 scenario: window B loaded the file before window A's save
    const path = await makePdfFile('duel.pdf')
    createPdfView(path)
    const windowA = lastWebContents.id
    createPdfView(path)
    const windowB = lastWebContents.id
    await openIn(windowA, path)
    await openIn(windowB, path)

    const first = await saveGranted(windowA, request(path))
    expect(first).toMatchObject({ ok: true })
    expect(showMessageBox).not.toHaveBeenCalled()

    // window B's baseline predates window A's write: its manual save must raise
    // the fence instead of silently overwriting (Overwrite → lands)
    showMessageBox.mockResolvedValue({ response: 1 })
    const second = await saveGranted(windowB, request(path))
    expect(second).toMatchObject({ ok: true })
    expect(showMessageBox).toHaveBeenCalledTimes(1)

    // ...and once overwritten, window B's baseline is fresh again
    showMessageBox.mockClear()
    const third = await saveGranted(windowB, request(path))
    expect(third).toMatchObject({ ok: true })
    expect(showMessageBox).not.toHaveBeenCalled()
  })

  it('an explicit Save As target is not fenced (user picked that exact file in a dialog)', async () => {
    const path = await makePdfFile('doc.pdf')
    const copyPath = join(path, '..', 'picked.pdf')
    createPdfView(path)
    const wcId = lastWebContents.id
    await openIn(wcId, path)

    const external = await pdfBytes(200)
    writeFileSync(path, external)

    // the fence's own Save As grants the picked copy (same grant the menu Save As uses)
    showMessageBox.mockResolvedValue({ response: 0 }) // Save As (default)
    showSaveDialog.mockResolvedValue({ canceled: false, filePath: copyPath })
    const fenced = await saveGranted(wcId, request(path))
    expect(fenced).toMatchObject({ ok: true, savedAsPath: copyPath })

    // a save addressed at the explicitly picked target skips the fence entirely:
    // only the first dialog was shown, and the copy is rewritten again
    const again = await saveGranted(wcId, request(path, { targetPath: copyPath }))
    expect(again).toMatchObject({ ok: true })
    expect(showMessageBox).toHaveBeenCalledTimes(1)
  })
})

/**
 * BUG-1654: the in-place save must verify that the file on disk still looks
 * like what this window last read/wrote (mtime+size stamp taken at open and
 * refreshed after every save). A blind write resurrects a renamed/deleted
 * path as a silent fork and lets a second window on the same file win
 * last-writer-wins without a warning. On a mismatch the handler must refuse
 * to write and offer Overwrite / Save As / Cancel; automatic (autosave) saves
 * are declined without a modal.
 */
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

type IpcHandler = (event: { sender: FakeWebContents }, ...args: unknown[]) => unknown

interface FakeWebContents {
  id: number
  listeners: Map<string, () => void>
  isDestroyed: ReturnType<typeof vi.fn>
  loadFile: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  once: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
}

const handlers = new Map<string, IpcHandler>()
const showMessageBox = vi.fn()
const showSaveDialogWithMemory = vi.fn()
const webContents: FakeWebContents[] = []
let nextWebContentsId = 1

function makeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const contents: FakeWebContents = {
    id: nextWebContentsId++,
    listeners,
    isDestroyed: vi.fn(() => false),
    loadFile: vi.fn(),
    loadURL: vi.fn(),
    once: vi.fn((event: string, listener: () => void) => listeners.set(event, listener)),
    send: vi.fn(),
    setWindowOpenHandler: vi.fn(),
  }
  webContents.push(contents)
  return contents
}

vi.mock('electron', () => ({
  app: {
    getPath: vi.fn(() => tmpdir()),
    on: vi.fn(),
    quit: vi.fn(),
    whenReady: vi.fn(() => new Promise(() => {})),
  },
  BrowserWindow: class {
    static fromWebContents() {
      return null
    }
    static getFocusedWindow() {
      return null
    }
  },
  dialog: {
    showMessageBox: (...args: unknown[]) => showMessageBox(...args),
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => handlers.set(channel, handler)),
    on: vi.fn((channel: string, handler: IpcHandler) => handlers.set(channel, handler)),
    removeHandler: vi.fn(),
  },
  net: { fetch: vi.fn() },
  protocol: { handle: vi.fn() },
  shell: { openExternal: vi.fn() },
  WebContentsView: class {
    webContents = makeWebContents()
  },
}))

vi.mock('@airy-office/electron-utils', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@airy-office/electron-utils')>()
  const { writeFile } = await import('node:fs/promises')
  return {
    ...actual,
    TextRecoveryStore: class {
      clear(): void {}
      async writeCopy(): Promise<void> {}
      async maybeRecover(_p: unknown, _r: unknown) {
        return { text: '', recovered: false }
      }
    },
    atomicWriteFile: async (path: string, data: Uint8Array) => void (await writeFile(path, data)),
    configuredDefaultSaveDir: vi.fn(() => tmpdir()),
    contextMenuLabels: vi.fn(() => ({})),
    installContextMenu: vi.fn(),
    installNavigationGuard: vi.fn(),
    safeExternalUrl: vi.fn(() => null),
    showOpenDialogWithMemory: vi.fn(),
    showSaveDialogWithMemory: (...args: unknown[]) => showSaveDialogWithMemory(...args),
  }
})

import { createMarkdownView, markdownFilePath } from '../src/main/markdown-main'
import { MARKDOWN_CHANNELS } from '../src/shared/ipc'

const temporaryDirectories: string[] = []

async function createDocument(content = '# Saved document'): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'markdown-save-staleness-'))
  temporaryDirectories.push(directory)
  const documentPath = join(directory, 'note.md')
  await writeFile(documentPath, content)
  return documentPath
}

function save(contents: FakeWebContents, request: Record<string, unknown>): Promise<unknown> {
  return handlers.get(MARKDOWN_CHANNELS.save)?.({ sender: contents }, request) as Promise<unknown>
}

/** Dialog button roles in promptExternalChange: [Save As, Overwrite, Cancel] */
const SAVE_AS = 0
const OVERWRITE = 1
const CANCEL = 2

afterEach(async () => {
  vi.restoreAllMocks()
  showMessageBox.mockReset()
  showSaveDialogWithMemory.mockReset()
  for (const contents of webContents.splice(0)) contents.listeners.get('destroyed')?.()
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('markdown save staleness fence', () => {
  it('saves an untouched file without prompting', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'save' })

    expect(result).toMatchObject({ ok: true, path: documentPath })
    expect(showMessageBox).not.toHaveBeenCalled()
    expect(await readFile(documentPath, 'utf8')).toBe('# Local edit')
  })

  it('refuses to write when the file was changed externally and the user cancels', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    await writeFile(documentPath, '# External edit')
    showMessageBox.mockResolvedValue({ response: CANCEL })

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'save' })

    expect(result).toMatchObject({ ok: true, canceled: true })
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    // the external version wins — no silent clobber
    expect(await readFile(documentPath, 'utf8')).toBe('# External edit')
    // the tab still points at the old path with the document still dirty:
    // honest UI state, the user keeps their unsaved edits
    expect(markdownFilePath(contents.id)).toBe(documentPath)
  })

  it('overwrites after a deliberate choice and refreshes the fence baseline', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    await writeFile(documentPath, '# External edit')
    showMessageBox.mockResolvedValue({ response: OVERWRITE })

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'save' })

    expect(result).toMatchObject({ ok: true, path: documentPath })
    expect(await readFile(documentPath, 'utf8')).toBe('# Local edit')

    // a follow-up save with no new external change must not prompt again —
    // the baseline now reflects the file as rewritten by this window
    showMessageBox.mockClear()
    const again = await save(contents, { text: '# Local edit 2', imageSources: [], mode: 'save' })
    expect(again).toMatchObject({ ok: true, path: documentPath })
    expect(showMessageBox).not.toHaveBeenCalled()
    expect(await readFile(documentPath, 'utf8')).toBe('# Local edit 2')
  })

  it('falls back to Save As after a refusal and leaves the external file alone', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    await writeFile(documentPath, '# External edit')
    showMessageBox.mockResolvedValue({ response: SAVE_AS })
    const targetPath = join(dirname(documentPath), 'my-version.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'save' })

    expect(result).toMatchObject({ ok: true, path: targetPath })
    expect(await readFile(targetPath, 'utf8')).toBe('# Local edit')
    expect(await readFile(documentPath, 'utf8')).toBe('# External edit')
    expect(markdownFilePath(contents.id)).toBe(targetPath)
  })

  it('does not resurrect the old path after an external rename', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    const movedPath = join(dirname(documentPath), 'moved.md')
    await rename(documentPath, movedPath)
    showMessageBox.mockResolvedValue({ response: CANCEL })

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'save' })

    expect(result).toMatchObject({ ok: true, canceled: true })
    expect(existsSync(documentPath)).toBe(false)
    expect(await readFile(movedPath, 'utf8')).toBe('# Saved document')
  })

  it('declines an automatic save on a stale file without a dialog', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    await writeFile(documentPath, '# External edit')

    const result = await save(contents, {
      text: '# Local edit',
      imageSources: [],
      mode: 'save',
      auto: true,
    })

    expect(result).toMatchObject({ ok: true, canceled: true })
    expect(showMessageBox).not.toHaveBeenCalled()
    expect(await readFile(documentPath, 'utf8')).toBe('# External edit')
  })

  it('fences the second window after the first window saved the shared file', async () => {
    const documentPath = await createDocument()
    const viewA = createMarkdownView(documentPath)
    const contentsA = viewA.webContents as unknown as FakeWebContents
    const viewB = createMarkdownView(documentPath)
    const contentsB = viewB.webContents as unknown as FakeWebContents

    // window A saves its edit — the on-disk file now differs from B's baseline
    const first = await save(contentsA, { text: '# From A', imageSources: [], mode: 'save' })
    expect(first).toMatchObject({ ok: true, path: documentPath })

    // window B (stale buffer) must be told, not silently win last-writer-wins
    showMessageBox.mockResolvedValue({ response: CANCEL })
    const second = await save(contentsB, { text: '# From B', imageSources: [], mode: 'save' })
    expect(second).toMatchObject({ ok: true, canceled: true })
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    expect(await readFile(documentPath, 'utf8')).toBe('# From A')
  })

  it('does not fence a Save As onto a freshly picked path', async () => {
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    const targetPath = join(dirname(documentPath), 'picked.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'saveAs' })

    expect(result).toMatchObject({ ok: true, path: targetPath })
    expect(showMessageBox).not.toHaveBeenCalled()
  })

  it('captures a baseline for an untitled first save, so a later external change is fenced', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'markdown-save-staleness-'))
    temporaryDirectories.push(directory)
    const view = createMarkdownView()
    const contents = view.webContents as unknown as FakeWebContents
    const targetPath = join(directory, 'first.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })

    const first = await save(contents, { text: '# First', imageSources: [], mode: 'saveAs' })
    expect(first).toMatchObject({ ok: true, path: targetPath })

    // the just-saved file was changed externally: the fence now applies
    await writeFile(targetPath, '# External edit')
    showMessageBox.mockResolvedValue({ response: CANCEL })
    const second = await save(contents, { text: '# Second', imageSources: [], mode: 'save' })
    expect(second).toMatchObject({ ok: true, canceled: true })
    expect(await readFile(targetPath, 'utf8')).toBe('# External edit')
  })

  it('keeps saving working after the fenced file reappears with a new directory', async () => {
    // regression guard for the asset paths: the fence must not break the
    // Save As asset relocation that follows a refused in-place save
    const documentPath = await createDocument()
    const view = createMarkdownView(documentPath)
    const contents = view.webContents as unknown as FakeWebContents
    const subdirectory = join(dirname(documentPath), 'elsewhere')
    await mkdir(subdirectory)
    await writeFile(documentPath, '# External edit')
    showMessageBox.mockResolvedValue({ response: SAVE_AS })
    const targetPath = join(subdirectory, 'note.md')
    showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })

    const result = await save(contents, { text: '# Local edit', imageSources: [], mode: 'save' })

    expect(result).toMatchObject({ ok: true, path: targetPath })
    expect(await readFile(targetPath, 'utf8')).toBe('# Local edit')
  })
})

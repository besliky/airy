/**
 * BUG-1724: the slides in-place save must verify that the file on disk still
 * looks like what this session last read/wrote (mtime+size stamp taken at open
 * and refreshed after every save). A blind write lets a second app instance on
 * the same deck (separate user-data dirs, so both legitimately coexist) win
 * last-writer-wins silently — the audit's MARKA/MARKB repro lost editor B's
 * edits without any dialog. On a mismatch the handler must refuse to write and
 * offer Save As / Overwrite / Cancel; automatic (autosave-tick) saves are
 * declined without a modal.
 *
 * Behavior harness in the style of the BUG-1654 twin
 * (apps/markdown/tests/save-staleness.test.ts): slides-main is an Electron
 * main module, so electron + the dialog-bearing electron-utils surfaces are
 * mocked while the real fence helpers (statFileStamp/checkSaveStaleness) and
 * the real pptx save pipeline run.
 */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

type IpcHandler = (event: { sender: FakeWebContents }, ...args: unknown[]) => unknown

interface FakeWebContents {
  id: number
  isDestroyed: () => boolean
  send: ReturnType<typeof vi.fn>
  executeJavaScript: () => Promise<null>
}

const h = vi.hoisted(() => {
  // set to a fresh scratch dir by scratchDir() before each test's opens
  const state = { userDataDir: '' }
  return {
    state,
    showMessageBox: vi.fn(),
    showSaveDialogWithMemory: vi.fn(),
  }
})

const handlers = new Map<string, IpcHandler>()
let nextWebContentsId = 1

function makeWebContents(): FakeWebContents {
  return {
    id: nextWebContentsId++,
    isDestroyed: () => false,
    send: vi.fn(),
    executeJavaScript: async () => null,
  }
}

vi.mock('electron', () => ({
  app: {
    getPath: vi.fn(() => h.state.userDataDir),
    on: vi.fn(),
    quit: vi.fn(),
    whenReady: vi.fn(() => Promise.resolve()),
  },
  BrowserWindow: class {
    static fromWebContents() {
      return null
    }
    static getFocusedWindow() {
      return null
    }
    static getAllWindows() {
      return []
    }
  },
  WebContentsView: class {
    webContents = makeWebContents()
  },
  clipboard: { writeText: vi.fn(), readText: vi.fn(() => '') },
  desktopCapturer: { getSources: vi.fn(async () => ({ sources: [] })) },
  dialog: {
    showMessageBox: (...args: unknown[]) => h.showMessageBox(...args),
    showMessageBoxSync: vi.fn(() => 0),
    showSaveDialog: vi.fn(),
    showOpenDialog: vi.fn(),
    showErrorBox: vi.fn(),
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => handlers.set(channel, handler)),
    on: vi.fn(),
    removeHandler: vi.fn(),
  },
  Menu: {
    buildFromTemplate: vi.fn(() => ({})),
    setApplicationMenu: vi.fn(),
    getApplicationMenu: vi.fn(() => null),
  },
  nativeImage: { createFromBuffer: vi.fn(), createFromPath: vi.fn(), createEmpty: vi.fn() },
  session: { defaultSession: { setDisplayMediaRequestHandler: vi.fn() }, fromPartition: vi.fn() },
  shell: { openExternal: vi.fn(), openPath: vi.fn(), showItemInFolder: vi.fn() },
  webContents: {
    getAllWebContents: vi.fn(() => []),
    fromId: vi.fn(() => null),
    fromFrame: vi.fn(() => null),
  },
}))

vi.mock('@airy-office/electron-utils', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@airy-office/electron-utils')>()
  return {
    ...actual,
    // the fence helpers (statFileStamp / checkSaveStaleness / saveAsSuggestion)
    // stay real — they are part of the behavior under test
    rendererMayReadPath: vi.fn(() => true),
    grantRendererFileAccess: vi.fn(),
    forgetRendererFileAccess: vi.fn(),
    forgetWitnessedDrops: vi.fn(),
    configuredDefaultSaveDir: vi.fn(() => h.state.userDataDir),
    configuredAuthorName: vi.fn(() => ''),
    showSaveDialogWithMemory: (...args: unknown[]) => h.showSaveDialogWithMemory(...args),
    showOpenDialogWithMemory: vi.fn(async () => ({ canceled: true, filePaths: [] })),
    showMessageBoxGuarded: vi.fn(),
    fetchRemoteImage: vi.fn(),
    installContextMenu: vi.fn(),
    installNavigationGuard: vi.fn(),
    openHelpUrl: vi.fn(),
    safeExternalUrl: vi.fn(() => null),
    appMenuLabels: vi.fn(() => ({})),
    contextMenuLabels: vi.fn(() => ({})),
    toggleDevToolsItem: vi.fn(() => ({})),
    voidLoad: vi.fn(),
  }
})

// System font metrics would need real platform fonts; the fence tests use a
// blank single-slide deck whose layout never measures text.
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
  resetFontRegistry: vi.fn(),
  registerEmbeddedFonts: vi.fn(() => false),
  exportFontFaces: vi.fn(() => []),
  listPrivateFontFaces: vi.fn(() => []),
  getPrivateFontData: vi.fn(() => null),
  familyAvailable: vi.fn(() => true),
  fontFileFamilies: vi.fn(() => new Map()),
  setUserFontDir: vi.fn(),
}))

// shaped-metrics imports the harfbuzz .wasm asset at module level (unresolvable
// under vitest); the fence tests don't shape text, so stub both entry points.
vi.mock('../src/main/shaped-metrics', () => ({
  shapedMetricsReady: vi.fn(async () => undefined),
  refineComplexWidths: vi.fn(async () => false),
}))

// the 30s autosave setInterval at slides-main module scope would keep the
// worker's event loop alive after the tests: capture it on the fake clock
vi.useFakeTimers()

const { registerSlidesIpc } = await import('../src/main/slides-main')
const { createBlankPptx, openPptx, savePptx } = await import('@airy-office/pptx-engine')

vi.useRealTimers()

const temporaryDirectories: string[] = []

async function scratchDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'slides-save-fence-'))
  temporaryDirectories.push(dir)
  h.state.userDataDir = dir
  return dir
}

/** A real one-slide deck on disk, written the same way the app saves. */
async function createDeck(): Promise<string> {
  const dir = await scratchDir()
  const deckPath = join(dir, 'deck.pptx')
  const bytes = await savePptx(await openPptx(await createBlankPptx()))
  await writeFile(deckPath, bytes)
  return deckPath
}

async function openDeck(path: string): Promise<FakeWebContents> {
  const contents = makeWebContents()
  const open = handlers.get('slides:open-path')!
  const result = (await open({ sender: contents }, path, 1280)) as { path?: string } | null
  expect(result).toMatchObject({ path })
  return contents
}

function save(contents: FakeWebContents, auto = false): Promise<unknown> {
  return handlers.get('slides:save')!({ sender: contents }, auto) as Promise<unknown>
}

/** Dialog button roles in the external-change prompt: [Save As, Overwrite, Cancel] */
const SAVE_AS = 0
const OVERWRITE = 1
const CANCEL = 2

afterEach(async () => {
  h.showMessageBox.mockReset()
  h.showSaveDialogWithMemory.mockReset()
  for (const dir of temporaryDirectories.splice(0)) {
    await rm(dir, { recursive: true, force: true })
  }
})

describe('slides save staleness fence (BUG-1724)', () => {
  it('saves an untouched deck without prompting', async () => {
    await scratchDir()
    registerSlidesIpc()
    const deckPath = await createDeck()
    const contents = await openDeck(deckPath)

    const result = await save(contents)

    expect(result).toMatchObject({ ok: true, path: deckPath })
    expect(h.showMessageBox).not.toHaveBeenCalled()
    // the saved package is a readable pptx (zip magic) written by the save path
    const bytes = await readFile(deckPath)
    expect(bytes.subarray(0, 2).toString()).toBe('PK')
  })

  it('refuses to write and prompts when the deck was changed by another instance', async () => {
    await scratchDir()
    registerSlidesIpc()
    const deckPath = await createDeck()
    const contents = await openDeck(deckPath)

    // the "second editor" of the audit's MARKA/MARKB repro: another instance
    // saved the same file after this session opened it
    await writeFile(deckPath, await savePptx(await openPptx(await createBlankPptx())))
    // touch the bytes so the stamp (mtime+size) no longer matches
    await writeFile(deckPath, Buffer.concat([await readFile(deckPath), Buffer.alloc(1)]))
    h.showMessageBox.mockResolvedValue({ response: CANCEL })

    const result = await save(contents)

    expect(h.showMessageBox).toHaveBeenCalledTimes(1)
    const options = h.showMessageBox.mock.calls[0][0] as {
      buttons: string[]
      defaultId: number
      cancelId: number
    }
    expect(options.buttons).toHaveLength(3)
    expect(options.defaultId).toBe(0)
    expect(options.cancelId).toBe(2)
    expect(result).toMatchObject({ ok: false, reason: 'external-modified' })
  })

  it('overwrites after a deliberate choice and refreshes the fence baseline', async () => {
    await scratchDir()
    registerSlidesIpc()
    const deckPath = await createDeck()
    const contents = await openDeck(deckPath)
    const external = Buffer.concat([await readFile(deckPath), Buffer.from('external')])
    await writeFile(deckPath, external)
    h.showMessageBox.mockResolvedValue({ response: OVERWRITE })

    const result = await save(contents)

    expect(result).toMatchObject({ ok: true, path: deckPath })
    // our write replaced the external contents
    expect(await readFile(deckPath)).not.toEqual(external)

    // a follow-up save with no new external change must not prompt again —
    // the baseline now reflects the file as rewritten by this session
    h.showMessageBox.mockClear()
    const again = await save(contents)
    expect(again).toMatchObject({ ok: true, path: deckPath })
    expect(h.showMessageBox).not.toHaveBeenCalled()
  })

  it('falls back to Save As after a refusal and leaves the external file alone', async () => {
    await scratchDir()
    registerSlidesIpc()
    const deckPath = await createDeck()
    const contents = await openDeck(deckPath)
    const external = Buffer.concat([await readFile(deckPath), Buffer.from('external')])
    await writeFile(deckPath, external)
    h.showMessageBox.mockResolvedValue({ response: SAVE_AS })
    const targetPath = join(dirname(deckPath), 'my-version.pptx')
    h.showSaveDialogWithMemory.mockResolvedValue({ canceled: false, filePath: targetPath })

    const result = await save(contents)

    expect(result).toMatchObject({ ok: true, path: targetPath })
    // the chosen target holds our deck (a valid pptx)…
    const written = await readFile(targetPath)
    expect(written.subarray(0, 2).toString()).toBe('PK')
    // …and the external version was never clobbered
    expect(await readFile(deckPath)).toEqual(external)

    // the session now tracks the new path: a follow-up save lands there quietly
    h.showMessageBox.mockClear()
    const again = await save(contents)
    expect(again).toMatchObject({ ok: true, path: targetPath })
    expect(h.showMessageBox).not.toHaveBeenCalled()
  })

  it('declines an automatic save on a stale deck without a dialog', async () => {
    await scratchDir()
    registerSlidesIpc()
    const deckPath = await createDeck()
    const contents = await openDeck(deckPath)
    const external = Buffer.concat([await readFile(deckPath), Buffer.from('external')])
    await writeFile(deckPath, external)

    const result = await save(contents, true)

    expect(h.showMessageBox).not.toHaveBeenCalled()
    expect(result).toMatchObject({ ok: false, reason: 'external-modified' })
    // the external version wins — no silent clobber from the autosave tick
    expect(await readFile(deckPath)).toEqual(external)
  })

  it('does not resurrect a deck whose path was renamed or deleted externally', async () => {
    await scratchDir()
    registerSlidesIpc()
    const deckPath = await createDeck()
    const contents = await openDeck(deckPath)
    const movedPath = join(dirname(deckPath), 'moved.pptx')
    const { rename } = await import('node:fs/promises')
    await rename(deckPath, movedPath)
    h.showMessageBox.mockResolvedValue({ response: CANCEL })

    const result = await save(contents)

    expect(result).toMatchObject({ ok: false, reason: 'external-modified' })
    expect(h.showMessageBox).toHaveBeenCalledTimes(1)
    // the moved file keeps the user's (external) bytes
    expect((await readFile(movedPath)).subarray(0, 2).toString()).toBe('PK')
  })

  it('an untitled session (no stamp) first-saves into the drafts folder unfenced', async () => {
    // a session whose baseline was never taken (new blank deck) must degrade
    // to the old unfenced behavior, and its first save must land somewhere
    await scratchDir()
    registerSlidesIpc()
    const contents = makeWebContents()
    const newBlank = handlers.get('slides:new-blank')!
    await newBlank({ sender: contents }, 1280)

    const result = (await save(contents, false)) as { ok: boolean; path?: string }

    expect(result.ok).toBe(true)
    expect(result.path).toBeTruthy()
    expect(result.path!.startsWith(h.state.userDataDir)).toBe(true)
    expect((await readFile(result.path!)).subarray(0, 2).toString()).toBe('PK')
    expect(h.showMessageBox).not.toHaveBeenCalled()
  })
})

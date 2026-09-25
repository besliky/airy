/**
 * BUG-1738 (html twin of the markdown regression): the crash-recovery store
 * must resolve its autosave dir at call time, never at module evaluation. The
 * shell redirects userData (app.setPath — AIRY_USER_DATA in test drivers,
 * "Airy Dev" for plain unpacked runs) only after html-main's top-level code
 * has already run, so an eager capture anchored every copy to the default
 * profile (~/.config/Airy): unpacked runs silently lost crash recovery
 * (ENOENT — the dir did not exist) or leaked copies into an installed Airy's
 * profile, where a later packaged launch would offer them back.
 *
 * The scenario simulates the real ordering: the module is imported while
 * userData still points at the pre-setPath default, the "shell" installs the
 * real profile, and only then do the IPC handlers run. The kill -9 + relaunch
 * leg re-imports the module fresh (new process state) against the same
 * profile and requires Restore to hand back the edits.
 */
import { existsSync, readdirSync } from 'node:fs'
import { mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
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
const ipcOn = vi.fn()
const showMessageBox = vi.fn(async () => ({ response: 0 }))
let nextWebContentsId = 1

function makeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  return {
    id: nextWebContentsId++,
    listeners,
    isDestroyed: vi.fn(() => false),
    loadFile: vi.fn(),
    loadURL: vi.fn(),
    once: vi.fn((event: string, listener: () => void) => listeners.set(event, listener)),
    send: vi.fn(),
    setWindowOpenHandler: vi.fn(),
  }
}

// shared with the electron mock factory below (hoisted)
const state = vi.hoisted(() => ({ userData: '' }))

vi.mock('electron', () => ({
  app: {
    // the fake shell: only `state.userData` answers as the installed profile
    getPath: (name: string) => (name === 'userData' ? state.userData : tmpdir()),
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
    on: vi.fn((channel: string, handler: IpcHandler) => ipcOn(channel, handler)),
    removeHandler: vi.fn(),
  },
  net: { fetch: vi.fn() },
  protocol: { handle: vi.fn() },
  shell: { openExternal: vi.fn() },
  webContents: { getAll: vi.fn(() => []) },
  WebContentsView: class {
    webContents = makeWebContents()
  },
}))

import { HTML_CHANNELS } from '../src/shared/ipc'

const temporaryDirectories: string[] = []

async function makeProfile(name: string): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), `recovery-${name}-`))
  temporaryDirectories.push(dir)
  return dir
}

async function makeDocument(): Promise<string> {
  const dir = await makeProfile('doc')
  const docPath = join(dir, 'page.html')
  await writeFile(docPath, '<!doctype html><p>seed</p>')
  // backdate the file so a copy written now is strictly newer (the offer rule)
  const past = new Date(Date.now() - 60_000)
  await utimes(docPath, past, past)
  return docPath
}

async function pushRecovery(sender: FakeWebContents, docPath: string, text: string): Promise<void> {
  await handlers.get(HTML_CHANNELS.writeRecovery)?.({ sender }, docPath, text)
}

function readViaIpc(sender: FakeWebContents, docPath: string): Promise<unknown> {
  return handlers.get(HTML_CHANNELS.readFile)?.({ sender }, docPath) as Promise<unknown>
}

afterEach(async () => {
  handlers.clear()
  ipcOn.mockClear()
  vi.resetModules()
  showMessageBox.mockClear()
  for (const dir of temporaryDirectories.splice(0)) await rm(dir, { recursive: true, force: true })
})

describe('html crash recovery under a late userData redirect (BUG-1738)', () => {
  it('lands copies in the profile the shell installed after import, and Restore survives a kill', async () => {
    // module evaluation happens while userData still points at the default
    // profile — exactly the state html-main is imported in by the shell
    const defaultProfile = await makeProfile('default-profile')
    state.userData = defaultProfile
    await import('../src/main/html-main')

    // the shell's app.setPath: unpacked runs redirect here, AIRY_USER_DATA drivers too
    const devProfile = await makeProfile('airy-user-data')
    state.userData = devProfile

    const docPath = await makeDocument()
    const htmlMain = await import('../src/main/html-main')
    const view = htmlMain.createHtmlView(docPath)
    await pushRecovery(view.webContents, docPath, '<!doctype html><p>DIRTY EDIT</p>')

    // the copy exists in the installed profile — and nothing leaked into the
    // default one (the pre-fix store wrote there and hit ENOENT or worse)
    expect(existsSync(join(devProfile, 'html-autosave'))).toBe(true)
    expect(readdirSync(join(devProfile, 'html-autosave'))).toHaveLength(1)
    expect(existsSync(join(defaultProfile, 'html-autosave'))).toBe(false)

    // kill -9 + relaunch: fresh module state, same profile — Restore returns
    // the unsaved edits the copy carried
    vi.resetModules()
    handlers.clear()
    state.userData = devProfile
    const relaunched = await import('../src/main/html-main')
    const reopened = relaunched.createHtmlView(docPath)
    const result = (await readViaIpc(reopened.webContents, docPath)) as {
      text: string
      recovered: boolean
    }
    expect(result.recovered).toBe(true)
    expect(result.text).toContain('DIRTY EDIT')
  })

  it('stays safe without any setPath — an isolated launch keeps using its own profile', async () => {
    // standalone mode never redirects userData: the store must still work
    // (creating the autosave dir itself) inside whatever profile the process has
    const isolatedProfile = await makeProfile('isolated-profile')
    state.userData = isolatedProfile
    const htmlMain = await import('../src/main/html-main')
    const docPath = await makeDocument()
    expect(existsSync(join(isolatedProfile, 'html-autosave'))).toBe(false)

    const view = htmlMain.createHtmlView(docPath)
    await pushRecovery(view.webContents, docPath, '<!doctype html><p>ISOLATED EDIT</p>')
    expect(readdirSync(join(isolatedProfile, 'html-autosave'))).toHaveLength(1)

    const result = (await readViaIpc(view.webContents, docPath)) as {
      text: string
      recovered: boolean
    }
    expect(result.recovered).toBe(true)
    expect(result.text).toContain('ISOLATED EDIT')
  })
})

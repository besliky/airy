/**
 * BUG-1738: the crash-recovery store must resolve its autosave dir at call
 * time, never at module evaluation. The shell redirects userData (app.setPath
 * — AIRY_USER_DATA in test drivers, "Airy Dev" for plain unpacked runs) only
 * after markdown-main's top-level code has already run, so an eager capture
 * anchored every copy to the default profile (~/.config/Airy): unpacked runs
 * silently lost crash recovery (ENOENT — the dir did not exist) or leaked
 * copies into an installed Airy's profile, where a later packaged launch
 * would offer them back as "Recovered version found".
 *
 * The two scenarios below simulate the real ordering: the module is imported
 * while userData still points at the pre-setPath default, then the "shell"
 * installs the real profile, and only then do the IPC handlers run. The
 * kill -9 + relaunch leg re-imports the module fresh (new process state)
 * against the same profile and requires Restore to hand back the edits.
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
    on: vi.fn(),
    removeHandler: vi.fn(),
  },
  net: { fetch: vi.fn() },
  protocol: { handle: vi.fn() },
  shell: { openExternal: vi.fn() },
  WebContentsView: class {
    webContents = makeWebContents()
  },
}))

import { MARKDOWN_CHANNELS } from '../src/shared/ipc'

const temporaryDirectories: string[] = []

async function makeProfile(name: string): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), `recovery-${name}-`))
  temporaryDirectories.push(dir)
  return dir
}

async function makeDocument(): Promise<string> {
  const dir = await makeProfile('doc')
  const docPath = join(dir, 'note.md')
  await writeFile(docPath, '# seed')
  // backdate the file so a copy written now is strictly newer (the offer rule)
  const past = new Date(Date.now() - 60_000)
  await utimes(docPath, past, past)
  return docPath
}

async function pushRecovery(sender: FakeWebContents, docPath: string, text: string): Promise<void> {
  await handlers.get(MARKDOWN_CHANNELS.writeRecovery)?.({ sender }, docPath, text)
}

function readViaIpc(sender: FakeWebContents, docPath: string): Promise<unknown> {
  return handlers.get(MARKDOWN_CHANNELS.readFile)?.({ sender }, docPath) as Promise<unknown>
}

afterEach(async () => {
  handlers.clear()
  vi.resetModules()
  showMessageBox.mockClear()
  for (const dir of temporaryDirectories.splice(0)) await rm(dir, { recursive: true, force: true })
})

describe('markdown crash recovery under a late userData redirect (BUG-1738)', () => {
  it('lands copies in the profile the shell installed after import, and Restore survives a kill', async () => {
    // module evaluation happens while userData still points at the default
    // profile — exactly the state markdown-main is imported in by the shell
    const defaultProfile = await makeProfile('default-profile')
    state.userData = defaultProfile
    await import('../src/main/markdown-main')

    // the shell's app.setPath: unpacked runs redirect here, AIRY_USER_DATA drivers too
    const devProfile = await makeProfile('airy-user-data')
    state.userData = devProfile

    const docPath = await makeDocument()
    const view = (await import('../src/main/markdown-main')).createMarkdownView(docPath)
    await pushRecovery(view.webContents, docPath, '# DIRTY EDIT')

    // the copy exists in the installed profile — and nothing leaked into the
    // default one (the pre-fix store wrote there and hit ENOENT or worse)
    expect(existsSync(join(devProfile, 'markdown-autosave'))).toBe(true)
    expect(readdirSync(join(devProfile, 'markdown-autosave'))).toHaveLength(1)
    expect(existsSync(join(defaultProfile, 'markdown-autosave'))).toBe(false)
    // kill -9 + relaunch: fresh module state, same profile — Restore returns
    // the unsaved edits the copy carried
    vi.resetModules()
    handlers.clear()
    state.userData = devProfile
    const relaunched = await import('../src/main/markdown-main')
    const reopened = relaunched.createMarkdownView(docPath)
    const result = (await readViaIpc(reopened.webContents, docPath)) as {
      text: string
      recovered: boolean
    }
    expect(result).toEqual({ text: '# DIRTY EDIT', recovered: true })
  })

  it('stays safe without any setPath — an isolated launch keeps using its own profile', async () => {
    // standalone mode never redirects userData: the store must still work
    // (creating the autosave dir itself) inside whatever profile the process has
    const isolatedProfile = await makeProfile('isolated-profile')
    state.userData = isolatedProfile
    const markdownMain = await import('../src/main/markdown-main')
    const docPath = await makeDocument()
    expect(existsSync(join(isolatedProfile, 'markdown-autosave'))).toBe(false)

    const view = markdownMain.createMarkdownView(docPath)
    await pushRecovery(view.webContents, docPath, '# ISOLATED EDIT')
    expect(readdirSync(join(isolatedProfile, 'markdown-autosave'))).toHaveLength(1)

    const result = (await readViaIpc(view.webContents, docPath)) as {
      text: string
      recovered: boolean
    }
    expect(result).toEqual({ text: '# ISOLATED EDIT', recovered: true })
  })
})

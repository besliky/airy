/**
 * BUG-1741: a save must respect the file's remembered encoding. The write
 * used to be `Buffer.from(text, 'utf8')` unconditionally: a windows-1251
 * file the user had pinned (UX-1653/1696) was silently transcoded to UTF-8
 * while the persisted pick stayed behind claiming the old charset — the next
 * open decoded the new bytes AS windows-1251 and produced mojibake, and the
 * status-bar picker masked the stale override with a session-local "Auto
 * detect". Now the save encodes into the pinned charset (the picker can read
 * the truth back through the read-only getEncoding channel), a plain UTF-8
 * file saves exactly as before, and a text the pinned charset cannot
 * represent falls back to a lossless UTF-8 write with the now-false pick
 * dropped, so the reopen auto-detects the file the way it actually is.
 */
import { readFileSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { decodeBytesAsEncoding, readRememberedFileEncoding } from '../src/main/encoding-memory'
import { MARKDOWN_CHANNELS } from '../src/shared/ipc'

const userData = vi.hoisted(() => ({ current: '' }))

type IpcHandler = (event: { sender: FakeWebContents }, ...args: unknown[]) => unknown

interface FakeWebContents {
  id: number
  listeners: Map<string, () => void>
  isDestroyed: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  once: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
}

const handlers = new Map<string, IpcHandler>()
const webContents: FakeWebContents[] = []
let nextWebContentsId = 1

function makeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const contents: FakeWebContents = {
    id: nextWebContentsId++,
    listeners,
    isDestroyed: vi.fn(() => false),
    send: vi.fn(),
    once: vi.fn((event: string, listener: () => void) => listeners.set(event, listener)),
    setWindowOpenHandler: vi.fn(),
  }
  webContents.push(contents)
  return contents
}

vi.mock('electron', () => ({
  app: {
    getPath: vi.fn((name: string) => (name === 'userData' ? userData.current : tmpdir())),
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
  dialog: { showMessageBox: vi.fn() },
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

// markdown-main must NOT be imported statically: the hoisted electron mock
// needs the real userData directory to exist first (set in beforeAll).

// "Привет, мир" in windows-1251 bytes (0xCF 0xF0 …) — not valid UTF-8
const PRIVET_1251 = Buffer.from([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2, 0x2c, 0x20, 0xec, 0xe8, 0xf0])

let markdownMain: typeof import('../src/main/markdown-main')
let settingsPath: string
const temporaryDirectories: string[] = []

async function makeFile(name: string, bytes: Buffer): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'markdown-save-encoding-'))
  temporaryDirectories.push(directory)
  const filePath = join(directory, name)
  await writeFile(filePath, bytes)
  return filePath
}

function callHandler(channel: string, sender: FakeWebContents, ...args: unknown[]): unknown {
  return handlers.get(channel)?.({ sender }, ...args)
}

const readFileFor = (filePath: string, sender: FakeWebContents) =>
  callHandler(MARKDOWN_CHANNELS.readFile, sender, filePath) as Promise<{
    text: string
    recovered: boolean
  }>

const saveFor = (
  sender: FakeWebContents,
  request: { text: string; imageSources: string[]; mode: 'save' },
) => callHandler(MARKDOWN_CHANNELS.save, sender, request) as Promise<{ ok: boolean; path?: string }>

const getEncodingFor = (filePath: string, sender: FakeWebContents) =>
  callHandler(MARKDOWN_CHANNELS.getEncoding, sender, filePath) as string

beforeAll(async () => {
  userData.current = await mkdtemp(join(tmpdir(), 'markdown-save-encoding-userdata-'))
  temporaryDirectories.push(userData.current)
  settingsPath = join(userData.current, 'app-settings.json')
  markdownMain = await import('../src/main/markdown-main')
})

afterEach(async () => {
  for (const contents of webContents.splice(0)) contents.listeners.get('destroyed')?.()
})

afterAll(async () => {
  await Promise.all(
    temporaryDirectories.map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('markdown save respects the remembered encoding', () => {
  it('writes a pinned windows-1251 file back as windows-1251; reopen is mojibake-free', async () => {
    const filePath = await makeFile('cyr.md', PRIVET_1251)
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    // open decodes through the pick once it is set
    await callHandler(MARKDOWN_CHANNELS.setEncoding, sender, filePath, 'windows-1251')
    expect(await readFileFor(filePath, sender)).toEqual({ text: 'Привет, мир', recovered: false })

    const edit = 'Привет, мир и до встречи'
    const result = await saveFor(sender, { text: edit, imageSources: [], mode: 'save' })
    expect(result).toMatchObject({ ok: true, path: filePath })

    // the disk bytes are the pinned charset, not the old unconditional UTF-8
    const onDisk = readFileSync(filePath)
    expect([...onDisk.slice(0, 2)]).toEqual([0xcf, 0xf0]) // 'П' in windows-1251
    expect(onDisk.equals(Buffer.from(edit, 'utf8'))).toBe(false)
    expect(decodeBytesAsEncoding(onDisk, 'windows-1251')).toBe(edit)

    // the pick survived (it still tells the truth) and the reopen reads the
    // edit without mojibake
    expect(getEncodingFor(filePath, sender)).toBe('windows-1251')
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBe('windows-1251')
    expect((await readFileFor(filePath, sender)).text).toBe(edit)
  })

  it('saves a plain UTF-8 file exactly as before, creating no pick', async () => {
    const filePath = await makeFile('plain.md', Buffer.from('# Заголовок\n', 'utf8'))
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    const result = await saveFor(sender, {
      text: '# Заголовок и текст\n',
      imageSources: [],
      mode: 'save',
    })
    expect(result).toMatchObject({ ok: true, path: filePath })
    expect(readFileSync(filePath).equals(Buffer.from('# Заголовок и текст\n', 'utf8'))).toBe(true)
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBeUndefined()
  })

  it('falls back to a lossless UTF-8 write and drops the pick when the charset cannot represent the text', async () => {
    const filePath = await makeFile('mixed.md', Buffer.from('hello\n', 'utf8'))
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents
    await callHandler(MARKDOWN_CHANNELS.setEncoding, sender, filePath, 'windows-1251')

    const result = await saveFor(sender, {
      text: '你好 world\n',
      imageSources: [],
      mode: 'save',
    })
    expect(result).toMatchObject({ ok: true, path: filePath })

    // CJK has no windows-1251 form: the write stays lossless as UTF-8 and the
    // now-false pick is gone, so the reopen auto-detects correctly
    expect(readFileSync(filePath).equals(Buffer.from('你好 world\n', 'utf8'))).toBe(true)
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect((await readFileFor(filePath, sender)).text).toBe('你好 world\n')
  })

  it('keeps a UTF-16le file and its BOM across save', async () => {
    const payload = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('Данные', 'utf16le')])
    const filePath = await makeFile('wide.md', payload)
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    // the BOM outranks any pick on decode; the kept-BOM text round-trips
    expect((await readFileFor(filePath, sender)).text).toBe('\uFEFFДанные')
    await callHandler(MARKDOWN_CHANNELS.setEncoding, sender, filePath, 'utf-16le')

    // the envelope re-emits the BOM, so the saved text carries it again
    const saved = '\uFEFFДанные!'
    const result = await saveFor(sender, { text: saved, imageSources: [], mode: 'save' })
    expect(result).toMatchObject({ ok: true })

    const onDisk = readFileSync(filePath)
    expect([...onDisk.slice(0, 2)]).toEqual([0xff, 0xfe])
    expect(decodeBytesAsEncoding(onDisk, 'utf-16le')).toBe(saved)
    expect(getEncodingFor(filePath, sender)).toBe('utf-16le')
    expect((await readFileFor(filePath, sender)).text).toBe(saved)
  })

  it('answers getEncoding only for granted paths', async () => {
    const filePath = await makeFile('granted.md', Buffer.from('x', 'utf8'))
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect(() => getEncodingFor('/elsewhere.md', sender)).toThrow('path not granted')
  })
})

/**
 * UX-1653: a manually chosen encoding must survive reopen. The pick is
 * persisted per file path in the shared workspace app-settings.json
 * (`fileEncodings`, most-recently-used first, capped at 200) through the
 * single-writer settings queue, and the open path (readFile handler →
 * TextRecoveryStore.readOriginal → readTextDecoded) decodes with the
 * remembered charset instead of re-running the very guess the user already
 * corrected (BUG-1646 / BUG-1651). A path without a pick keeps the plain
 * auto-detection behavior.
 */
import { mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import {
  decodeBytesAsEncoding,
  parseFileEncodings,
  readRememberedFileEncoding,
  recordFileEncoding,
  rememberFileEncoding,
} from '../src/main/encoding-memory'
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

// markdown-main must NOT be imported statically: its module body seeds the
// recovery store from userData at load time, and the hoisted electron mock
// needs the real userData directory to exist first (set in beforeAll).

// "Привет, мир" in windows-1251 bytes (0xCF 0xF0 …) — not valid UTF-8
const PRIVET_1251 = Buffer.from([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2, 0x2c, 0x20, 0xec, 0xe8, 0xf0])

let markdownMain: typeof import('../src/main/markdown-main')
let settingsPath: string
const temporaryDirectories: string[] = []

async function makeFile(name: string, bytes: Buffer): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'markdown-encoding-'))
  temporaryDirectories.push(directory)
  const filePath = join(directory, name)
  await writeFile(filePath, bytes)
  return filePath
}

function readFileFor(filePath: string, sender: FakeWebContents): Promise<{ text: string }> {
  return handlers.get(MARKDOWN_CHANNELS.readFile)?.({ sender }, filePath) as Promise<{
    text: string
  }>
}

function pickEncoding(
  filePath: string,
  encoding: string | null,
  sender: FakeWebContents,
): Promise<boolean> {
  // UX-1696: the channel moved into the shared registry once the preload
  // grew its pass-through; null is the "Auto" pick (forget, back to detect)
  return handlers.get(MARKDOWN_CHANNELS.setEncoding)?.(
    { sender },
    filePath,
    encoding,
  ) as Promise<boolean>
}

beforeAll(async () => {
  // must exist before markdown-main initializes (recoveryStore seeds its
  // autosave dir from userData at import time)
  userData.current = mkdtempSync(join(tmpdir(), 'markdown-encoding-userdata-'))
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

// ── the persisted LRU (pure part) ──

describe('recordFileEncoding', () => {
  it('adds a new path at the front', () => {
    const entries = recordFileEncoding([], '/a.md', 'gb18030')
    expect(entries).toEqual([{ path: '/a.md', encoding: 'gb18030' }])
  })

  it('moves a re-picked path to the front with the new charset', () => {
    const entries = recordFileEncoding(
      [
        { path: '/a.md', encoding: 'gb18030' },
        { path: '/b.md', encoding: 'koi8-r' },
      ],
      '/a.md',
      'windows-1251',
    )
    expect(entries).toEqual([
      { path: '/a.md', encoding: 'windows-1251' },
      { path: '/b.md', encoding: 'koi8-r' },
    ])
  })

  it('caps the list at 200 entries, evicting the least recently used', () => {
    let entries: ReturnType<typeof recordFileEncoding> = []
    for (let i = 0; i < 201; i++) entries = recordFileEncoding(entries, `/f${i}.md`, 'utf-8')
    expect(entries).toHaveLength(200)
    expect(entries[0]).toEqual({ path: '/f200.md', encoding: 'utf-8' })
    // /f0.md — the oldest pick — fell off the end
    expect(entries.some((entry) => entry.path === '/f0.md')).toBe(false)
    expect(entries.some((entry) => entry.path === '/f1.md')).toBe(true)
  })

  it('honors a custom cap', () => {
    // persisted order is most-recently-used first
    let entries = [
      { path: '/newer.md', encoding: 'utf-8' },
      { path: '/oldest.md', encoding: 'utf-8' },
    ]
    entries = recordFileEncoding(entries, '/newest.md', 'utf-8', 2)
    expect(entries).toEqual([
      { path: '/newest.md', encoding: 'utf-8' },
      { path: '/newer.md', encoding: 'utf-8' },
    ])
  })
})

describe('parseFileEncodings', () => {
  it('returns an empty list for anything that is not an array', () => {
    expect(parseFileEncodings(undefined)).toEqual([])
    expect(parseFileEncodings({ path: '/a.md', encoding: 'utf-8' })).toEqual([])
  })

  it('drops malformed entries but keeps valid ones', () => {
    expect(
      parseFileEncodings([
        { path: '/ok.md', encoding: 'gb18030' },
        { path: '/no-encoding.md' },
        { path: '', encoding: 'utf-8' },
        { encoding: 'utf-8' },
        { path: '/bad-charset.md', encoding: 'iso-9001' },
        { path: '/num.md', encoding: 42 },
        'garbage',
        null,
      ]),
    ).toEqual([{ path: '/ok.md', encoding: 'gb18030' }])
  })
})

// ── persistence through the shared single-writer settings queue ──

describe('rememberFileEncoding persistence', () => {
  it('round-trips a pick and preserves unrelated settings keys', async () => {
    writeFileSync(
      settingsPath,
      JSON.stringify({ language: 'zh', lastDialogDirs: [{ scope: 'markdown', dir: '/d' }] }),
    )
    await rememberFileEncoding(settingsPath, '/a.md', 'windows-1251')
    expect(readRememberedFileEncoding(settingsPath, '/a.md')).toBe('windows-1251')
    const settings = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    // the merge-write kept keys this feature never touches (OBS-1532 discipline)
    expect(settings.language).toBe('zh')
    expect(settings.lastDialogDirs).toEqual([{ scope: 'markdown', dir: '/d' }])
  })

  it('keeps concurrent picks and foreign writes from dropping each other', async () => {
    const { queueAppSettingsUpdate } = await import('@airy-office/electron-utils')
    await Promise.all([
      rememberFileEncoding(settingsPath, '/a.md', 'gb18030'),
      rememberFileEncoding(settingsPath, '/b.md', 'shift_jis'),
      queueAppSettingsUpdate(settingsPath, { language: 'en' }),
      rememberFileEncoding(settingsPath, '/c.md', 'big5'),
    ])
    expect(readRememberedFileEncoding(settingsPath, '/a.md')).toBe('gb18030')
    expect(readRememberedFileEncoding(settingsPath, '/b.md')).toBe('shift_jis')
    expect(readRememberedFileEncoding(settingsPath, '/c.md')).toBe('big5')
    const settings = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect(settings.language).toBe('en')
    // atomic temp+rename: the file always parses and no staging file remains
    expect(readdirSync(userData.current).filter((name) => name.endsWith('.tmp'))).toEqual([])
  })
})

// ── the forced decode (engine-parity semantics) ──

describe('decodeBytesAsEncoding', () => {
  it('decodes exactly as the remembered charset says', () => {
    expect(decodeBytesAsEncoding(PRIVET_1251, 'windows-1251')).toBe('Привет, мир')
  })

  it('keeps a BOM authoritative over the pick, byte-identical on round-trip', () => {
    const withBom = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('héllo', 'utf8')])
    // the kept-BOM decode matches the editors' open path: the U+FEFF stays in
    // the text so an untouched open→save round-trips byte-identically
    expect(decodeBytesAsEncoding(withBom, 'windows-1251')).toBe('\uFEFFhéllo')
    expect(withBom[0]).toBe(0xef)
  })

  it('handles a UTF-16 BOM file regardless of the pick', () => {
    const bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('你好', 'utf16le')])
    expect(decodeBytesAsEncoding(bytes, 'windows-1251')).toBe('\uFEFF你好')
  })

  it('forces utf-8 honestly: invalid bytes surface as replacement characters', () => {
    expect(decodeBytesAsEncoding(PRIVET_1251, 'utf-8')).toContain('\uFFFD')
  })
})

// ── the main-process wiring: pick → reopen applies it ──

describe('markdown reopen with remembered encoding', () => {
  it('applies the remembered pick when the same path reopens', async () => {
    const filePath = await makeFile('note.md', Buffer.from('héllo wörld', 'utf8'))
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    // before the pick, the file decodes with the normal auto-detection
    const before = await readFileFor(filePath, sender)
    expect(before).toEqual({ text: 'héllo wörld', recovered: false })

    await pickEncoding(filePath, 'windows-1251', sender)

    // the pick is authoritative: the reopen decodes AS windows-1251, even
    // against the engine's unconditional strict-UTF-8 win
    const after = await readFileFor(filePath, sender)
    expect(after.text).toBe('hГ©llo wГ¶rld')
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBe('windows-1251')
  })

  it('keeps a non-UTF-8 file correct across reopen once its charset is picked', async () => {
    const filePath = await makeFile('ru.md', PRIVET_1251)
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    await pickEncoding(filePath, 'windows-1251', sender)

    const first = await readFileFor(filePath, sender)
    const second = await readFileFor(filePath, sender)
    expect(first.text).toBe('Привет, мир')
    expect(second.text).toBe('Привет, мир')
  })

  it('auto-detects a different file without a pick exactly as before', async () => {
    const pickedPath = await makeFile('picked.md', Buffer.from('héllo wörld', 'utf8'))
    const plainPath = await makeFile('plain.md', Buffer.from('Русский текст', 'utf8'))
    const pickedView = markdownMain.createMarkdownView(pickedPath)
    await pickEncoding(
      pickedPath,
      'windows-1251',
      pickedView.webContents as unknown as FakeWebContents,
    )

    // another tab, another file, never picked: plain auto-detection
    const plainView = markdownMain.createMarkdownView(plainPath)
    const result = await readFileFor(plainPath, plainView.webContents as unknown as FakeWebContents)
    expect(result.text).toBe('Русский текст')
    expect(readRememberedFileEncoding(settingsPath, plainPath)).toBeUndefined()
  })

  it('drops the pick again when the Auto option forgets it (UX-1696)', async () => {
    const filePath = await makeFile('auto.md', PRIVET_1251)
    const view = markdownMain.createMarkdownView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    await pickEncoding(filePath, 'windows-1251', sender)
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBe('windows-1251')

    await pickEncoding(filePath, null, sender)
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBeUndefined()

    // without the pick the open path is back to plain auto-detection
    const reopened = await readFileFor(filePath, sender)
    expect(reopened.text).toBe('Привет, мир')
  })

  it('rejects ungranted paths and unknown charsets without persisting anything', async () => {
    const grantedPath = await makeFile('granted.md', Buffer.from('x', 'utf8'))
    const view = markdownMain.createMarkdownView(grantedPath)
    const sender = view.webContents as unknown as FakeWebContents

    await expect(pickEncoding('/elsewhere.md', 'gb18030', sender)).rejects.toThrow(
      'path not granted',
    )
    await expect(pickEncoding(grantedPath, 'iso-9001', sender)).rejects.toThrow('unknown encoding')
    expect(readRememberedFileEncoding(settingsPath, '/elsewhere.md')).toBeUndefined()
    expect(readRememberedFileEncoding(settingsPath, grantedPath)).toBeUndefined()
  })
})

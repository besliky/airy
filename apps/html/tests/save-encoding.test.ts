/**
 * BUG-1782 (the html half of BUG-1741/#237): a save must respect the file's
 * remembered encoding and the saved file's <meta charset> claim must match
 * the bytes on disk. The write used to be `Buffer.from(text, 'utf8')`
 * unconditionally with the declaration left untouched: a windows-1251 file
 * the user had pinned (UX-1653/1696) was silently transcoded to UTF-8 while
 * both the persisted pick and the meta claim kept saying windows-1251 — the
 * next open (and every browser) decoded the new bytes AS windows-1251 and
 * produced mojibake. Now the save encodes into the pinned charset (readable
 * back through the read-only getEncoding channel), a text the pinned charset
 * cannot represent falls back to a lossless UTF-8 write with the now-false
 * pick dropped, the declaration is synced to the charset the bytes actually
 * use (both directions — a legacy-declared file edited with no pick re-encodes
 * to UTF-8 and its claim is rewritten, BUG-761 parity), and back-to-back
 * saves of a pinned file are byte-idempotent.
 */
import { readFileSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { decodeBytesAsEncoding, readRememberedFileEncoding } from '../src/main/encoding-memory'
import { encodeTextAsEncoding } from '@airy-office/electron-utils'
import { HTML_CHANNELS } from '../src/shared/ipc'

const userData = vi.hoisted(() => ({ current: '' }))

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
  webContents: { getAll: vi.fn(() => []) },
  WebContentsView: class {
    webContents = makeWebContents()
  },
}))

// html-main must NOT be imported statically: the hoisted electron mock
// needs the real userData directory to exist first (set in beforeAll).

/** the template the cp1251 fixtures are encoded from (ASCII skeleton, Cyrillic content) */
const PAGE_1251 = [
  '<!doctype html>',
  '<html>',
  '<head><meta charset="windows-1251"><title>Кодировка</title></head>',
  '<body><h1>Привет, мир</h1></body>',
  '</html>',
  '',
].join('\n')

const EDIT_1251 = PAGE_1251.replace('<h1>Привет, мир</h1>', '<h1>Привет, мир и до встречи</h1>')

/** the declaration token as the main process syncs it (the only rewrite a save makes) */
const withClaim = (text: string, charset: string): string =>
  text.replace(/(<meta charset=")[^"]*(")/, `$1${charset}$2`)

let htmlMain: typeof import('../src/main/html-main')
let settingsPath: string
const temporaryDirectories: string[] = []

async function makeFile(name: string, bytes: Buffer): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), 'html-save-encoding-'))
  temporaryDirectories.push(directory)
  const filePath = join(directory, name)
  await writeFile(filePath, bytes)
  return filePath
}

function callHandler(channel: string, sender: FakeWebContents, ...args: unknown[]): unknown {
  return handlers.get(channel)?.({ sender }, ...args)
}

const readFileFor = (filePath: string, sender: FakeWebContents) =>
  callHandler(HTML_CHANNELS.readFile, sender, filePath) as Promise<{
    text: string
    recovered: boolean
  }>

const saveFor = (
  sender: FakeWebContents,
  request: { text: string; imageSources: string[]; mode: 'save' },
) =>
  callHandler(HTML_CHANNELS.save, sender, request) as Promise<{
    ok: boolean
    path?: string
  }>

const getEncodingFor = (filePath: string, sender: FakeWebContents) =>
  callHandler(HTML_CHANNELS.getEncoding, sender, filePath) as string | null

beforeAll(async () => {
  userData.current = await mkdtemp(join(tmpdir(), 'html-save-encoding-userdata-'))
  temporaryDirectories.push(userData.current)
  settingsPath = join(userData.current, 'app-settings.json')
  htmlMain = await import('../src/main/html-main')
})

afterEach(async () => {
  for (const contents of webContents.splice(0)) contents.listeners.get('destroyed')?.()
})

afterAll(async () => {
  await Promise.all(
    temporaryDirectories.map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('html save respects the remembered encoding', () => {
  it('writes a pinned windows-1251 file back as windows-1251; two cycles are byte-idempotent', async () => {
    const filePath = await makeFile('cyr.html', encodeTextAsEncoding(PAGE_1251, 'windows-1251')!)
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    // open decodes through the pick once it is set
    await callHandler(HTML_CHANNELS.setEncoding, sender, filePath, 'windows-1251')
    expect((await readFileFor(filePath, sender)).text).toBe(PAGE_1251)

    const first = await saveFor(sender, { text: EDIT_1251, imageSources: [], mode: 'save' })
    expect(first).toMatchObject({ ok: true, path: filePath })

    // the disk bytes are the pinned charset — not the old unconditional UTF-8 —
    // and the meta claim still names the encoding the bytes actually use
    const onDisk = readFileSync(filePath)
    expect(onDisk.equals(Buffer.from(EDIT_1251, 'utf8'))).toBe(false)
    expect(decodeBytesAsEncoding(onDisk, 'windows-1251')).toBe(EDIT_1251)
    expect(onDisk.toString('latin1')).toContain('<meta charset="windows-1251">')

    // second save of the same text: identical bytes, no drift
    const second = await saveFor(sender, { text: EDIT_1251, imageSources: [], mode: 'save' })
    expect(second).toMatchObject({ ok: true })
    expect(readFileSync(filePath).equals(onDisk)).toBe(true)

    // the pick survived (it still tells the truth) and the reopen reads the
    // edit without mojibake
    expect(getEncodingFor(filePath, sender)).toBe('windows-1251')
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBe('windows-1251')
    expect((await readFileFor(filePath, sender)).text).toBe(EDIT_1251)
  })

  it('rewrites a stale utf-8 claim when the pick pins windows-1251', async () => {
    // the seeded file's bytes are cp1251 while its claim says utf-8 — the
    // renderer buffer keeps that lying claim, and the save must fix it
    const lying = PAGE_1251.replace('<meta charset="windows-1251">', '<meta charset="utf-8">')
    const filePath = await makeFile('liar.html', encodeTextAsEncoding(lying, 'windows-1251')!)
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents
    await callHandler(HTML_CHANNELS.setEncoding, sender, filePath, 'windows-1251')
    expect((await readFileFor(filePath, sender)).text).toBe(lying)

    const edited = lying.replace('<h1>Привет, мир</h1>', '<h1>Привет, мир и до встречи</h1>')
    const result = await saveFor(sender, { text: edited, imageSources: [], mode: 'save' })
    expect(result).toMatchObject({ ok: true })

    // the bytes are cp1251, so the honest claim is windows-1251 — browsers
    // trusting the stale utf-8 label would render every Cyrillic run as mojibake
    const onDisk = readFileSync(filePath)
    const honest = withClaim(edited, 'windows-1251')
    expect(decodeBytesAsEncoding(onDisk, 'windows-1251')).toBe(honest)
    expect(decodeBytesAsEncoding(onDisk, 'windows-1251')).toContain('<meta charset="windows-1251">')
    expect((await readFileFor(filePath, sender)).text).toBe(honest)
  })

  it('re-encodes a legacy-declared file to UTF-8 on a no-pick edited save and rewrites the claim', async () => {
    // the audit's damage case WITHOUT any picker involvement: decode by the
    // declared charset, save UTF-8, declaration untouched — browsers then
    // decoded the new bytes as the old charset
    const filePath = await makeFile('legacy.html', encodeTextAsEncoding(PAGE_1251, 'windows-1251')!)
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    const result = await saveFor(sender, { text: EDIT_1251, imageSources: [], mode: 'save' })
    expect(result).toMatchObject({ ok: true })

    const savedUtf8 = withClaim(EDIT_1251, 'utf-8')
    const onDisk = readFileSync(filePath)
    expect(onDisk.equals(Buffer.from(savedUtf8, 'utf8'))).toBe(true)
    expect(onDisk.toString('utf8')).toContain('<meta charset="utf-8">')
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect((await readFileFor(filePath, sender)).text).toBe(savedUtf8)
  })

  it('falls back to a lossless UTF-8 write, drops the pick and fixes the claim when the charset cannot represent the text', async () => {
    const filePath = await makeFile('mixed.html', encodeTextAsEncoding(PAGE_1251, 'windows-1251')!)
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents
    await callHandler(HTML_CHANNELS.setEncoding, sender, filePath, 'windows-1251')

    const edited = EDIT_1251.replace('</h1>', ' 你好</h1>')
    const result = await saveFor(sender, { text: edited, imageSources: [], mode: 'save' })
    expect(result).toMatchObject({ ok: true })

    // CJK has no windows-1251 form: the write stays lossless as UTF-8, the
    // now-false pick is gone and the declaration no longer claims 1251, so
    // both the reopen and any browser auto-detect the file the way it is
    const savedUtf8 = withClaim(edited, 'utf-8')
    const onDisk = readFileSync(filePath)
    expect(onDisk.equals(Buffer.from(savedUtf8, 'utf8'))).toBe(true)
    expect(onDisk.toString('utf8')).toContain('<meta charset="utf-8">')
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect((await readFileFor(filePath, sender)).text).toBe(savedUtf8)
  })

  it('keeps gb18030 two-byte text pinned; a four-byte-only character falls back honestly', async () => {
    const two = encodeTextAsEncoding(
      '<!doctype html>\n<html>\n<head><meta charset="gb18030"></head>\n<body><p>你好</p></body>\n</html>\n',
      'gb18030',
    )!
    const filePath = await makeFile('cjk.html', two)
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents
    await callHandler(HTML_CHANNELS.setEncoding, sender, filePath, 'gb18030')

    // 你好 live in the two-byte grid: the save stays gb18030, claim intact
    const edited =
      '<!doctype html>\n<html>\n<head><meta charset="gb18030"></head>\n<body><p>你好，世界</p></body>\n</html>\n'
    const first = await saveFor(sender, { text: edited, imageSources: [], mode: 'save' })
    expect(first).toMatchObject({ ok: true })
    const onDisk = readFileSync(filePath)
    expect(decodeBytesAsEncoding(onDisk, 'gb18030')).toBe(edited)
    expect(onDisk.toString('latin1')).toContain('<meta charset="gb18030">')
    expect(getEncodingFor(filePath, sender)).toBe('gb18030')

    // 𠮷 (U+20BB7) is reachable only through the four-byte form the inverse
    // map does not model: an honest lossless UTF-8 fallback, pick dropped,
    // claim rewritten — the next open auto-detects instead of misreading
    const astral = withClaim(edited.replace('你好，世界', '你好 𠮷'), 'utf-8')
    const second = await saveFor(sender, {
      text: edited.replace('你好，世界', '你好 𠮷'),
      imageSources: [],
      mode: 'save',
    })
    expect(second).toMatchObject({ ok: true })
    const fallback = readFileSync(filePath)
    expect(fallback.equals(Buffer.from(astral, 'utf8'))).toBe(true)
    expect(fallback.toString('utf8')).toContain('<meta charset="utf-8">')
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect((await readFileFor(filePath, sender)).text).toBe(astral)
  })

  it('saves a plain UTF-8 file exactly as before, creating no pick', async () => {
    const filePath = await makeFile('plain.html', Buffer.from('<p>Заголовок</p>\n', 'utf8'))
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents

    const result = await saveFor(sender, {
      text: '<p>Заголовок и текст</p>\n',
      imageSources: [],
      mode: 'save',
    })
    expect(result).toMatchObject({ ok: true, path: filePath })
    expect(readFileSync(filePath).equals(Buffer.from('<p>Заголовок и текст</p>\n', 'utf8'))).toBe(
      true,
    )
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect(readRememberedFileEncoding(settingsPath, filePath)).toBeUndefined()
  })

  it('answers getEncoding only for granted paths', async () => {
    const filePath = await makeFile('granted.html', Buffer.from('<p>x</p>', 'utf8'))
    const view = htmlMain.createHtmlView(filePath)
    const sender = view.webContents as unknown as FakeWebContents
    expect(getEncodingFor(filePath, sender)).toBeNull()
    expect(() => getEncodingFor('/elsewhere.html', sender)).toThrow('path not granted')
  })
})

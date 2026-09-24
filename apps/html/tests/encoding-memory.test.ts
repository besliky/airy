/**
 * UX-1653: a manually chosen encoding must survive reopen. The html app
 * shares the markdown app's mechanism: the pick is persisted per file path
 * in the shared workspace app-settings.json (`fileEncodings`, LRU, cap 200)
 * through the single-writer settings queue, and the open path decodes with
 * the remembered charset — outranking both the declared <meta charset> and
 * the detector (BUG-1651) — while any path without a pick keeps the plain
 * auto-detection behavior.
 *
 * The html-main wiring is pinned by source assertions (the repo's pattern
 * for this app, cf. save-staleness-wiring.test.ts); the mechanics — LRU,
 * persistence, forced decode — are behavior-tested below against this
 * app's own copy of encoding-memory.
 */
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'

import {
  decodeBytesAsEncoding,
  parseFileEncodings,
  readRememberedFileEncoding,
  recordFileEncoding,
  rememberFileEncoding,
  removeFileEncoding,
} from '../src/main/encoding-memory'

const here = dirname(fileURLToPath(import.meta.url))
const mainSource = readFileSync(join(here, '../src/main/html-main.ts'), 'utf8')

function sourceContains(needle: string): boolean {
  return mainSource.includes(needle)
}

// "Привет, мир" in windows-1251 bytes (0xCF 0xF0 …) — not valid UTF-8
const PRIVET_1251 = Buffer.from([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2, 0x2c, 0x20, 0xec, 0xe8, 0xf0])

const cleanup: string[] = []

function settingsPathIn(name: string): string {
  const dir = mkdtempSync(join(tmpdir(), 'html-encoding-'))
  cleanup.push(dir)
  return join(dir, name)
}

afterEach(() => {
  for (const dir of cleanup.splice(0)) rmSync(dir, { recursive: true, force: true })
})

// ── the persisted LRU (pure part) ──

describe('recordFileEncoding', () => {
  it('adds a new path at the front', () => {
    const entries = recordFileEncoding([], '/a.html', 'gb18030')
    expect(entries).toEqual([{ path: '/a.html', encoding: 'gb18030' }])
  })

  it('moves a re-picked path to the front with the new charset', () => {
    const entries = recordFileEncoding(
      [
        { path: '/a.html', encoding: 'gb18030' },
        { path: '/b.html', encoding: 'koi8-r' },
      ],
      '/a.html',
      'windows-1251',
    )
    expect(entries).toEqual([
      { path: '/a.html', encoding: 'windows-1251' },
      { path: '/b.html', encoding: 'koi8-r' },
    ])
  })

  it('caps the list at 200 entries, evicting the least recently used', () => {
    let entries: ReturnType<typeof recordFileEncoding> = []
    for (let i = 0; i < 201; i++) entries = recordFileEncoding(entries, `/f${i}.html`, 'utf-8')
    expect(entries).toHaveLength(200)
    expect(entries[0]).toEqual({ path: '/f200.html', encoding: 'utf-8' })
    // /f0.html — the oldest pick — fell off the end
    expect(entries.some((entry) => entry.path === '/f0.html')).toBe(false)
    expect(entries.some((entry) => entry.path === '/f1.html')).toBe(true)
  })
})

describe('removeFileEncoding', () => {
  it('drops the path pick and keeps the rest of the LRU (the Auto option)', () => {
    const entries = [
      { path: '/a.html', encoding: 'windows-1251' },
      { path: '/b.html', encoding: 'koi8-r' },
    ]
    expect(removeFileEncoding(entries, '/a.html')).toEqual([
      { path: '/b.html', encoding: 'koi8-r' },
    ])
    // forgetting an unknown path is a no-op
    expect(removeFileEncoding(entries, '/missing.html')).toEqual(entries)
  })
})

describe('parseFileEncodings', () => {
  it('returns an empty list for anything that is not an array', () => {
    expect(parseFileEncodings(undefined)).toEqual([])
    expect(parseFileEncodings('nope')).toEqual([])
  })

  it('drops malformed entries but keeps valid ones', () => {
    expect(
      parseFileEncodings([
        { path: '/ok.html', encoding: 'big5' },
        { path: '/bad-charset.html', encoding: 'iso-9001' },
        { path: '', encoding: 'utf-8' },
        { path: '/num.html', encoding: 42 },
        'garbage',
        null,
      ]),
    ).toEqual([{ path: '/ok.html', encoding: 'big5' }])
  })
})

// ── persistence through the shared single-writer settings queue ──

describe('rememberFileEncoding persistence', () => {
  it('round-trips a pick and preserves unrelated settings keys', async () => {
    const settingsPath = settingsPathIn('app-settings.json')
    writeFileSync(
      settingsPath,
      JSON.stringify({ language: 'zh', lastDialogDirs: [{ scope: 'html', dir: '/d' }] }),
    )
    await rememberFileEncoding(settingsPath, '/a.html', 'shift_jis')
    expect(readRememberedFileEncoding(settingsPath, '/a.html')).toBe('shift_jis')
    const settings = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    // the merge-write kept keys this feature never touches (OBS-1532 discipline)
    expect(settings.language).toBe('zh')
    expect(settings.lastDialogDirs).toEqual([{ scope: 'html', dir: '/d' }])
  })

  it('keeps concurrent picks and foreign writes from dropping each other', async () => {
    const settingsPath = settingsPathIn('app-settings.json')
    const { queueAppSettingsUpdate } = await import('@airy-office/electron-utils')
    await Promise.all([
      rememberFileEncoding(settingsPath, '/a.html', 'gb18030'),
      rememberFileEncoding(settingsPath, '/b.html', 'windows-1251'),
      queueAppSettingsUpdate(settingsPath, { language: 'en' }),
      rememberFileEncoding(settingsPath, '/c.html', 'big5'),
    ])
    expect(readRememberedFileEncoding(settingsPath, '/a.html')).toBe('gb18030')
    expect(readRememberedFileEncoding(settingsPath, '/b.html')).toBe('windows-1251')
    expect(readRememberedFileEncoding(settingsPath, '/c.html')).toBe('big5')
    const settings = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect(settings.language).toBe('en')
    // atomic temp+rename: the file always parses and no staging file remains
    const dir = dirname(settingsPath)
    expect(readdirSync(dir).filter((name) => name.endsWith('.tmp'))).toEqual([])
  })
})

// ── the forced decode: the pick outranks the declared charset ──

describe('decodeBytesAsEncoding', () => {
  it('decodes exactly as the remembered charset says', () => {
    expect(decodeBytesAsEncoding(PRIVET_1251, 'windows-1251')).toBe('Привет, мир')
  })

  it('keeps a BOM authoritative over the pick', () => {
    const withBom = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('héllo', 'utf8')])
    // the kept-BOM decode matches the editors' open path (U+FEFF stays in the text)
    expect(decodeBytesAsEncoding(withBom, 'windows-1251')).toBe('\uFEFFhéllo')
  })

  it('forces utf-8 honestly: invalid bytes surface as replacement characters', () => {
    expect(decodeBytesAsEncoding(PRIVET_1251, 'utf-8')).toContain('\uFFFD')
  })
})

// ── the html-main wiring (source pins, cf. save-staleness-wiring.test.ts) ──

describe('html-main encoding-memory wiring', () => {
  it('the open path decodes with a remembered pick before declared/detected charsets', () => {
    expect(sourceContains('const remembered = readRememberedFileEncoding(appSettingsPath(), path)'))
    expect(sourceContains('if (remembered) return decodeBytesAsEncoding(bytes, remembered)'))
  })

  it('the set-encoding channel fences the path, validates the charset and forgets on null', () => {
    // UX-1696 moved the channel into the pinned HTML_CHANNELS registry (the
    // preload passes it through); null is the "Auto" pick (back to detect)
    expect(sourceContains('HTML_CHANNELS.setEncoding,'))
    expect(sourceContains('html: path not granted to this view'))
    expect(sourceContains('if (encoding === null)'))
    expect(sourceContains('await forgetFileEncoding(appSettingsPath(), path)'))
    expect(sourceContains('if (!isSelectableEncoding(encoding))'))
    expect(sourceContains('rememberFileEncoding(appSettingsPath(), path, encoding)'))
  })
})

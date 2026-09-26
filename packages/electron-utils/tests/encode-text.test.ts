/**
 * BUG-1741: the save path must be able to write the charset a file's
 * encoding-memory pick names, not only UTF-8. Node ships no legacy
 * TextEncoder, so encode-text builds each charset's inverse decode table
 * from its TextDecoder. These tests pin exact well-known byte forms (the
 * same pairs the decoder itself produced in verification), ASCII passthrough
 * for every selectable charset, decode-encode-decode round-trips, and the
 * honest null when a text cannot live in the target charset.
 */
import { describe, expect, it } from 'vitest'

import { encodeTextAsEncoding } from '../src/encode-text'

/**
 * The charsets the editors let a user pick (SELECTABLE_ENCODINGS in both
 * apps' shared/ipc — the package must not depend on app code). Keep in sync
 * with that list; the ASCII-passthrough pin below walks exactly it.
 */
const SELECTABLE_ENCODINGS = [
  'utf-8',
  'utf-16le',
  'utf-16be',
  'gb18030',
  'shift_jis',
  'big5',
  'euc-kr',
  'windows-1252',
  'windows-1251',
  'koi8-r',
  'windows-1250',
  'windows-1253',
  'windows-1255',
  'windows-1256',
  'windows-874',
] as const

const decode = (charset: string, bytes: Uint8Array): string =>
  new TextDecoder(charset).decode(bytes)

describe('encodeTextAsEncoding single-byte charsets', () => {
  it('encodes the exact windows-1251 byte form of Cyrillic', () => {
    // "Привет" — the same byte list encoding-memory.test.ts decodes with
    expect([...encodeTextAsEncoding('Привет', 'windows-1251')!]).toEqual([
      0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2,
    ])
  })

  it('encodes the exact byte form of other single-byte alphabets', () => {
    expect([...encodeTextAsEncoding('é', 'windows-1252')!]).toEqual([0xe9])
    expect([...encodeTextAsEncoding('Щ', 'koi8-r')!]).toEqual([0xfd])
    expect(decode('windows-1250', encodeTextAsEncoding('Ř', 'windows-1250')!)).toBe('Ř')
    expect(decode('windows-1253', encodeTextAsEncoding('αβ', 'windows-1253')!)).toBe('αβ')
    expect(decode('windows-1255', encodeTextAsEncoding('אב', 'windows-1255')!)).toBe('אב')
    expect(decode('windows-1256', encodeTextAsEncoding('سلام', 'windows-1256')!)).toBe('سلام')
    expect(decode('windows-874', encodeTextAsEncoding('ก', 'windows-874')!)).toBe('ก')
  })

  it('keeps ASCII byte-identical in utf-8 and every single-byte charset', () => {
    for (const charset of SELECTABLE_ENCODINGS) {
      if (charset === 'utf-16le' || charset === 'utf-16be') continue // 2 bytes per char by design
      const ascii = encodeTextAsEncoding('Header: 42\n', charset)!
      expect(ascii.equals(Buffer.from('Header: 42\n', 'utf8'))).toBe(true)
    }
  })

  it('returns null for a character the charset cannot represent', () => {
    expect(encodeTextAsEncoding('你好', 'windows-1251')).toBeNull()
    // an unknown charset has no inverse map at all
    expect(encodeTextAsEncoding('Привет', 'iso-9001')).toBeNull()
  })

  it('returns null for a lone surrogate instead of writing mojibake', () => {
    expect(encodeTextAsEncoding('\uDEAD', 'windows-1251')).toBeNull()
  })
})

describe('encodeTextAsEncoding double-byte charsets', () => {
  it('encodes the exact well-known two-byte forms', () => {
    expect([...encodeTextAsEncoding('你', 'gb18030')!]).toEqual([0xc4, 0xe3])
    expect([...encodeTextAsEncoding('あ', 'shift_jis')!]).toEqual([0x82, 0xa0])
    expect([...encodeTextAsEncoding('中', 'big5')!]).toEqual([0xa4, 0xa4])
    expect([...encodeTextAsEncoding('한', 'euc-kr')!]).toEqual([0xc7, 0xd1])
  })

  it('encodes shift_jis half-width katakana as single bytes', () => {
    expect([...encodeTextAsEncoding('ｱ', 'shift_jis')!]).toEqual([0xb1])
  })

  it('round-trips representative text through the real decoder', () => {
    const samples: Array<[string, string]> = [
      ['gb18030', '简体中文，测试文本。'],
      ['shift_jis', 'ひらがな、カタカナと漢字。ｱｲｳ'],
      ['big5', '繁體中文，測試文字。'],
      ['euc-kr', '한국어 텍스트입니다.'],
      ['windows-1251', 'Привет, мир — кириллица жива'],
      ['koi8-r', 'Русский текст'],
      ['windows-1252', 'accented: éàüñç'],
    ]
    for (const [charset, text] of samples) {
      expect(decode(charset, encodeTextAsEncoding(text, charset)!)).toBe(text)
    }
  })
})

describe('encodeTextAsEncoding UTF family', () => {
  it('encodes utf-8 exactly like Buffer', () => {
    expect(encodeTextAsEncoding('Привет 你好', 'utf-8')).toEqual(Buffer.from('Привет 你好', 'utf8'))
  })

  it('encodes utf-16le and utf-16be with the right byte order', () => {
    expect([...encodeTextAsEncoding('a', 'utf-16le')!]).toEqual([0x61, 0x00])
    expect([...encodeTextAsEncoding('A', 'utf-16le')!]).toEqual([0x41, 0x00])
    expect([...encodeTextAsEncoding('A', 'utf-16be')!]).toEqual([0x00, 0x41])
    expect([...encodeTextAsEncoding('Ж', 'utf-16be')!]).toEqual([0x04, 0x16])
    expect(decode('utf-16be', encodeTextAsEncoding('Привет 你好', 'utf-16be')!)).toBe('Привет 你好')
  })

  it('encodes an empty string to zero bytes without failing', () => {
    expect(encodeTextAsEncoding('', 'windows-1251')).toHaveLength(0)
    expect(encodeTextAsEncoding('', 'utf-8')).toHaveLength(0)
  })
})

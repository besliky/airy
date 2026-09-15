import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { decodeHtmlText, decodeTextBytes, legacyCharsetForLang } from '../src/text'
import { writeFixture } from './helpers/fixtures'

/** First two lines of a real windows-1251 stock CSV (Cyrillic header row). */
const CP1251_CSV_HEX =
  '22c0f0f2e8eaf3eb223b22cde0e8ece5edeee2e0ede8e5223b22cfe0f0e0ece5f2f0fb22' +
  '3b22caeeebe8f7e5f1f2e2ee223b22532f4e223b22a8eceaeef1f2fc223b22c7e0efeeeb' +
  'ede5ededeef1f2fc220d0a22313532303235343439223b22c7e5edeae5f02031352c3878' +
  '3230204c313034205a33223b22d3efe0eaeee2eae03a20312c20d1eef1f2eeffede8e53a' +
  '20cdeee2fbe9223b393b3b31303b39300d0a'
const CP1251_CSV_TEXT =
  '"Артикул";"Наименование";"Параметры";"Количество";"S/N";"Ёмкость";"Заполненность"\r\n' +
  '"152025449";"Зенкер 15,8x20 L104 Z3";"Упаковка: 1, Состояние: Новый";9;;10;90\r\n'

/**
 * Node has no legacy encoder, so build the byte string by brute-forcing each
 * character against the decoder: every code point below 0x80 is itself, the
 * rest is found by scanning the charset's byte range.
 */
function encodeLegacy(text: string, charset: string, maxLead = 0xff, trailRange = 0): Buffer {
  const decoder = new TextDecoder(charset)
  const out: number[] = []
  for (const character of text) {
    const code = character.codePointAt(0)!
    if (code < 0x80) {
      out.push(code)
      continue
    }
    let found = false
    if (trailRange === 0) {
      for (let b = 0x80; b <= maxLead && !found; b += 1) {
        if (decoder.decode(new Uint8Array([b])) === character) {
          out.push(b)
          found = true
        }
      }
    } else {
      for (let lead = 0x81; lead <= maxLead && !found; lead += 1) {
        for (let trail = trailRange; trail <= 0xfe && !found; trail += 1) {
          if (decoder.decode(new Uint8Array([lead, trail])) === character) {
            out.push(lead, trail)
            found = true
          }
        }
      }
    }
    if (!found) throw new Error(`cannot encode ${character} in ${charset}`)
  }
  return Buffer.from(out)
}

const cp1251 = (text: string): Buffer => encodeLegacy(text, 'windows-1251')
const cp1252 = (text: string): Buffer => encodeLegacy(text, 'windows-1252')
const cp1250 = (text: string): Buffer => encodeLegacy(text, 'windows-1250')
const gbk = (text: string): Buffer => encodeLegacy(text, 'gb18030', 0xfe, 0x40)
const shiftJis = (text: string): Buffer => encodeLegacy(text, 'shift_jis', 0xfc, 0x40)

describe('decodeTextBytes', () => {
  const text = 'город,житель\n'

  it('reads UTF-8 without a BOM', () => {
    expect(decodeTextBytes(Buffer.from(text, 'utf8'))).toBe(text)
  })

  it('reads UTF-8 and UTF-16 BOMs as a leading U+FEFF character, like readFile(utf8)', () => {
    // editors write that character back as BOM bytes on save, so an
    // untouched open→save round-trip must stay byte-identical
    expect(
      decodeTextBytes(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, 'utf8')])),
    ).toBe(`﻿${text}`)
    expect(decodeTextBytes(Buffer.from(`﻿${text}`, 'utf16le'))).toBe(`﻿${text}`)
    expect(decodeTextBytes(Buffer.from(`﻿${text}`, 'utf16le').swap16())).toBe(`﻿${text}`)
  })

  it('keeps plain ASCII untouched', () => {
    expect(decodeTextBytes(Buffer.from('a,b\n1,2\n', 'utf8'))).toBe('a,b\n1,2\n')
  })

  it('decodes a real windows-1251 CSV without any language hint', () => {
    expect(decodeTextBytes(Buffer.from(CP1251_CSV_HEX, 'hex'))).toBe(CP1251_CSV_TEXT)
  })

  it('decodes windows-1251 with the ru hint', () => {
    const text = 'Обзор рынка;цена,руб\n'
    expect(decodeTextBytes(cp1251(text), 'windows-1251')).toBe(text)
  })

  it('does not read a windows-1252 file as Cyrillic', () => {
    const text = 'Produit;Prix TTC\nRésumé très bien après le déjeuner\n'
    expect(decodeTextBytes(cp1252(text))).toBe(text)
    // the Russian hint must not flip it either: a locale guess never
    // overrides a script the file barely contains
    expect(decodeTextBytes(cp1252(text), 'windows-1251')).toBe(text)
  })

  it('decodes windows-1250 Polish, with and without the pl hint', () => {
    const text = 'Zażółć gęślą jaźń\nKraków, Wrocław, Łódź\n'
    expect(decodeTextBytes(cp1250(text), 'windows-1250')).toBe(text)
    expect(decodeTextBytes(cp1250(text))).toBe(text)
  })

  it('decodes GBK Chinese without a hint and Shift_JIS with the ja hint', () => {
    const zh = '城市,人口\n'
    const ja = '都市,人口\n東京,37\n'
    expect(decodeTextBytes(gbk(zh))).toBe(zh)
    expect(decodeTextBytes(shiftJis(ja), 'shift_jis')).toBe(ja)
    // GBK/Shift_JIS stay mutually ambiguous without a hint — the bytes decode
    // to plausible CJK under both, so nothing should confidently pick either
    expect(decodeTextBytes(shiftJis(ja))).not.toBe(ja)
  })
})

describe('legacyCharsetForLang', () => {
  it('maps UI languages to the legacy charset their locale writes', () => {
    expect(legacyCharsetForLang('ru')).toBe('windows-1251')
    expect(legacyCharsetForLang('pl')).toBe('windows-1250')
    expect(legacyCharsetForLang('he')).toBe('windows-1255')
    expect(legacyCharsetForLang('zh')).toBe('gb18030')
    expect(legacyCharsetForLang('zh-TW')).toBe('big5')
    expect(legacyCharsetForLang('fr')).toBe('windows-1252')
  })

  it('returns nothing for languages without a legacy charset', () => {
    expect(legacyCharsetForLang('hi')).toBeUndefined()
    expect(legacyCharsetForLang(undefined)).toBeUndefined()
  })
})

describe('decodeHtmlText', () => {
  it('trusts the declared meta charset over detection', () => {
    // windows-1252 bytes, but the page honestly declares windows-1251
    const text = '<html><head><meta charset="windows-1251"></head><body>Прайс</body></html>'
    expect(decodeHtmlText(cp1251(text))).toBe(text)
  })

  it('falls through to detection when a utf-8 claim does not hold', () => {
    const html =
      '<html><head><meta charset="utf-8"></head><body>Прайс-лист металлорежущего инструмента</body></html>'
    expect(decodeHtmlText(cp1251(html))).toBe(html)
  })

  it('detects an undeclared legacy page', () => {
    const text = '<html><body>Прайс-лист</body></html>'
    expect(decodeHtmlText(cp1251(text))).toBe(text)
  })
})

describe('parseFileToText: legacy-encoded attachments', () => {
  it('decodes a windows-1251 .txt attachment', async () => {
    const path = writeFixture('notes-cp1251.txt', cp1251('Заметки: проверить склад'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toBe('Заметки: проверить склад')
  })

  it('decodes a windows-1251 .csv attachment', async () => {
    const path = writeFixture('stock-cp1251.csv', Buffer.from(CP1251_CSV_HEX, 'hex'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('Артикул')
    expect(result.text).not.toContain('\uFFFD')
  })

  it('decodes a legacy .html attachment via its meta charset', async () => {
    const body =
      '<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1251"></head><body>Каталог</body></html>'
    const path = writeFixture('page-cp1251.html', cp1251(body))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('Каталог')
  })
})

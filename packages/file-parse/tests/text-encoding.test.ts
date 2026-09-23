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
    ).toBe(`\uFEFF${text}`)
    expect(decodeTextBytes(Buffer.from(`\uFEFF${text}`, 'utf16le'))).toBe(`\uFEFF${text}`)
    expect(decodeTextBytes(Buffer.from(`\uFEFF${text}`, 'utf16le').swap16())).toBe(`\uFEFF${text}`)
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

describe('decodeTextBytes: GBK at sentence length (BUG-1646)', () => {
  // windows-1250 soup used to win non-monotonically: scores flip on which
  // GBK trail bytes happen to land in ASCII range and which lead bytes hit
  // frequent Latin letters, so 8B read fine while 12B and 58B turned to mush.
  // Sentence-length coverage for every shape: tiny phrases, full sentences,
  // GBK-extension rows whose trail bytes are ASCII letters, and text mixed
  // with real ASCII (digits, URLs, CSV separators).
  const gbkVectors = [
    '你好世界', // 8 bytes
    '中文测试', // 8 bytes
    '城市,人口\n', // 10 bytes, ASCII separators
    '测试ABC数据', // 11 bytes, ASCII letters inside
    '中文测试语句', // 12 bytes
    '汉字编码标准', // 12 bytes
    '中文測試：這是通用規範漢字的一句話。第二行：辦公室自動化。', // 58 bytes, the audit sample
    '這是一段繁體中文的文字，用來測試編碼偵測的功能與正確性。', // 56 bytes, GBK-extension rows
    '中华人民共和国成立于一九四九年，是一个拥有悠久历史和灿烂文化的国家。', // 68 bytes
    '订单号:2024-0311,客户:张三,金额:1280.50元\n订单号:2024-0312,客户:李四,金额:980.00元\n',
    '城市,人口\n北京,2154\n上海,2428\n广州,1868\n深圳,1756\n杭州,1036\n',
    '欢迎访问 http://www.example.com 主页,如需帮助请联系客服。\n',
    '在日常办公中，电子表格软件被广泛用于数据统计和财务分析。用户可以通过公式快速计算总和、平均值以及增长率，从而提高工作效率并减少人工错误的发生。', // 142 bytes
    '这家公司成立于二十世纪九十年代，最初只是一家小型贸易办事处。经过三十年的发展，它已经成为一家拥有数千名员工、业务遍布全国各地的综合性企业集团，涉及制造、物流、金融服务等多个领域，并且在海外多个国家设有分支机构。', // 210 bytes
  ]

  it.each(gbkVectors)('decodes GBK text without a hint: %s', (text) => {
    expect(decodeTextBytes(gbk(text))).toBe(text)
  })

  it('decodes the audit sample with the zh hint too', () => {
    const text = '中文測試：這是通用規範漢字的一句話。第二行：辦公室自動化。'
    expect(decodeTextBytes(gbk(text), 'gb18030')).toBe(text)
  })

  // the same sentence shapes in the single-byte charsets must keep winning:
  // the soup tests must never cost a genuine file its script bonus
  const controls: Array<[string, Buffer, string]> = [
    [
      'cp1250 Polish',
      cp1250('Zażółć gęślą jaźń\nKraków, Wrocław, Łódź\n'),
      'Zażółć gęślą jaźń\nKraków, Wrocław, Łódź\n',
    ],
    [
      'cp1250 Czech sentence',
      cp1250(
        'Příliš žluťoučký kůň úpěl ďábelské ódy. Nástroj pro automatizované vyhledávání dokumentů byl otevřen ve čtvrtek.\n',
      ),
      'Příliš žluťoučký kůň úpěl ďábelské ódy. Nástroj pro automatizované vyhledávání dokumentů byl otevřen ve čtvrtek.\n',
    ],
    [
      'cp1251 Russian sentence',
      cp1251(
        'Город Москва,житель\nОбзор рынка металлорежущего инструмента за первый квартал года.\n',
      ),
      'Город Москва,житель\nОбзор рынка металлорежущего инструмента за первый квартал года.\n',
    ],
    [
      'cp1251 Russian CSV',
      cp1251('"Артикул";"Наименование";"Количество";"Цена"\r\n"152025449";"Зенкер";9;10\r\n'),
      '"Артикул";"Наименование";"Количество";"Цена"\r\n"152025449";"Зенкер";9;10\r\n',
    ],
    [
      'cp1252 French sentence',
      cp1252(
        'Résumé très bien après le déjeuner à Paris, où les étudiants français étaient réunis.\n',
      ),
      'Résumé très bien après le déjeuner à Paris, où les étudiants français étaient réunis.\n',
    ],
    [
      'cp1252 German sentence',
      cp1252('Müller kam spät aus Köln zurück, denn die Züge fuhren nicht mehr nach Fürth.\n'),
      'Müller kam spät aus Köln zurück, denn die Züge fuhren nicht mehr nach Fürth.\n',
    ],
    [
      'cp1252 German caps header',
      cp1252('GRÖSSE;ÄPFEL;SAFT\r\n12;3;4\r\n'),
      'GRÖSSE;ÄPFEL;SAFT\r\n12;3;4\r\n',
    ],
  ]

  it.each(controls)('still decodes %s at sentence length', (_label, bytes, text) => {
    expect(decodeTextBytes(bytes)).toBe(text)
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

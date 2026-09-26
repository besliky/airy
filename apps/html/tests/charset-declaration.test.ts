/**
 * BUG-1782: the saved file's <meta charset> declaration must name the charset
 * the bytes actually use — browsers trust the declaration, so a stale claim
 * renders the saved file as mojibake (the GUI editor used to leave the
 * declaration untouched while re-encoding the document; the MCP server fixed
 * its own saves the same way in BUG-761). These pins cover the declaration
 * forms browsers accept, canonical-alias and case comparisons (an alias like
 * cp1251 must not be churned by a windows-1251 save — untouched saves stay
 * byte-identical), and the untouched-ness of everything around the token.
 */
import { describe, expect, it } from 'vitest'

import { declaredHtmlCharset, syncCharsetDeclaration } from '../src/main/charset-declaration'

describe('syncCharsetDeclaration', () => {
  it('rewrites the charset attribute form in place', () => {
    const before = '<!doctype html><html><head><meta charset="windows-1251"><p>x'
    const { text, replaced } = syncCharsetDeclaration(before, 'utf-8')
    expect(replaced).toBe('windows-1251')
    expect(text).toBe('<!doctype html><html><head><meta charset="utf-8"><p>x')
  })

  it('rewrites the http-equiv content-type form, leaving the rest of the tag intact', () => {
    const before = '<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">'
    const { text, replaced } = syncCharsetDeclaration(before, 'utf-8')
    expect(replaced).toBe('windows-1251')
    expect(text).toBe('<meta http-equiv="Content-Type" content="text/html; charset=utf-8">')
  })

  it('rewrites an unquoted attribute value', () => {
    const { text } = syncCharsetDeclaration('<meta charset=windows-1251>', 'utf-8')
    expect(text).toBe('<meta charset=utf-8>')
  })

  it('keeps a declaration that already names the encoding byte-identical', () => {
    const before = '<meta charset="windows-1251">'
    const { text, replaced } = syncCharsetDeclaration(before, 'windows-1251')
    expect(text).toBe(before)
    expect(replaced).toBeNull()
  })

  it('compares canonically, so aliases are not churned', () => {
    // cp1251 is a WHATWG alias of windows-1251 — same encoding, no rewrite
    const cp = '<meta charset="cp1251">'
    expect(syncCharsetDeclaration(cp, 'windows-1251')).toEqual({ text: cp, replaced: null })
    // utf8 is an alias of utf-8
    const utf8 = '<meta charset="utf8">'
    expect(syncCharsetDeclaration(utf8, 'utf-8')).toEqual({ text: utf8, replaced: null })
  })

  it('compares case-insensitively', () => {
    const mixed = '<meta charset="Windows-1251">'
    expect(syncCharsetDeclaration(mixed, 'windows-1251')).toEqual({
      text: mixed,
      replaced: null,
    })
  })

  it('leaves a document without any declaration untouched', () => {
    const before = '<!doctype html><html><body><p>plain</p></body></html>'
    expect(syncCharsetDeclaration(before, 'windows-1251')).toEqual({
      text: before,
      replaced: null,
    })
  })

  it('rewrites a label no decoder accepts instead of trusting it', () => {
    const { text, replaced } = syncCharsetDeclaration('<meta charset="iso-9001">', 'utf-8')
    expect(replaced).toBe('iso-9001')
    expect(text).toBe('<meta charset="utf-8">')
  })

  it('splices only the token: CRLF line endings and neighboring tags survive', () => {
    const before = '<html>\r\n<head>\r\n<meta charset="windows-1251">\r\n<title>t</title>\r\n'
    const { text } = syncCharsetDeclaration(before, 'windows-1251')
    expect(text).toBe(before)
    const rewritten = syncCharsetDeclaration(before, 'utf-8').text
    expect(rewritten).toBe('<html>\r\n<head>\r\n<meta charset="utf-8">\r\n<title>t</title>\r\n')
  })

  it('rewrites the first declaration when several are present', () => {
    const before = '<meta charset="windows-1251"><meta charset="koi8-r">'
    const { text, replaced } = syncCharsetDeclaration(before, 'utf-8')
    expect(replaced).toBe('windows-1251')
    expect(text).toBe('<meta charset="utf-8"><meta charset="koi8-r">')
  })
})

describe('declaredHtmlCharset', () => {
  it('returns the token exactly as written', () => {
    expect(declaredHtmlCharset('<meta charset="Windows-1251">')).toBe('Windows-1251')
    expect(
      declaredHtmlCharset('<meta http-equiv="content-type" content="text/html; charset=koi8-r">'),
    ).toBe('koi8-r')
  })

  it('returns null when nothing is declared', () => {
    expect(declaredHtmlCharset('<p>no head</p>')).toBeNull()
  })
})

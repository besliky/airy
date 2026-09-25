import { describe, expect, it } from 'vitest'
import { sanitizeClipboardHtml } from '../src/renderer/editor/pasteSanitize'

// UX-1705: clipboard HTML from Word/web is sanitized before it reaches the
// schema parse. Policy: comments and dangerous elements removed, everything
// off the allow-list unwrapped (inline styles CUT, not simplified), attributes
// stripped down to a tiny semantic set with http(s)/mailto-only URLs.

// shape of a real Word clipboard fragment (mso styles, data islands,
// StartFragment markers, span wrappers, namespace-prefixed office tags)
const WORD_FRAGMENT = [
  '<html xmlns:o="urn:schemas-microsoft-com:office:office">',
  '<head><style><!--p.MsoNormal {mso-style-parent:""; font-family:"Arial";}--></style>',
  '<!--[if gte mso 9]><xml><w:WordDocument><w:View>Normal</w:View></w:WordDocument></xml><![endif]-->',
  '</head><body>',
  '<!--StartFragment-->',
  "<p class=MsoNormal style='mso-margin-top-alt:auto'><span style='font-family:\"Arial\"'>",
  'Hello <b>bold</b> <i>italic</i> <u>underline</u>',
  '</span><o:p></o:p></p>',
  '<h1>Heading <st1:place>One</st1:place></h1>',
  '<ul><li><span>first</span></li><li>second</li></ul>',
  '<table border=1><tr><td>a</td><td>b</td></tr></table>',
  '<!--EndFragment-->',
  '</body></html>',
].join('\n')

describe('sanitized Word fragments', () => {
  const out = sanitizeClipboardHtml(WORD_FRAGMENT)

  it('cuts mso styles, data islands and comment markers', () => {
    expect(out).not.toContain('mso')
    expect(out).not.toContain('<style')
    expect(out).not.toContain('WordDocument')
    expect(out).not.toContain('<!--')
    expect(out).not.toContain('StartFragment')
    expect(out).not.toContain('EndFragment')
  })

  it('keeps basic formatting: p, b/i/u, h1, li, table', () => {
    expect(out).toContain('<p>')
    expect(out).toContain('<b>bold</b>')
    expect(out).toContain('<i>italic</i>')
    expect(out).toContain('<u>underline</u>')
    expect(out).toContain('<h1>')
    expect(out).toContain('<ul>')
    expect(out).toContain('<li>first</li>')
    expect(out).toContain('<table>')
    expect(out).toContain('<td>a</td>')
    // the visible text survived wholesale
    expect(out).toContain('Hello ')
    expect(out).toContain('Heading ')
    expect(out).toContain('One')
  })

  it('unwraps span/font wrappers and namespace-prefixed office tags', () => {
    expect(out).not.toContain('<span')
    expect(out).not.toContain('<font')
    expect(out).not.toContain('<o:p')
    expect(out).not.toContain('<st1:')
    expect(out).not.toContain('style=')
  })
})

describe('dangerous content', () => {
  it('removes script/style/iframe with their subtree and on* handlers', () => {
    const out = sanitizeClipboardHtml(
      '<p onclick="evil()">a</p><script>alert(1)</script>' +
        '<style>.x{}</style><iframe src="https://e.c"></iframe><p>tail</p>',
    )
    expect(out).not.toContain('script')
    expect(out).not.toContain('alert')
    expect(out).not.toContain('iframe')
    expect(out).not.toContain('.x{')
    expect(out).not.toContain('onclick')
    expect(out).toContain('a')
    expect(out).toContain('tail')
  })

  it('keeps http(s)/mailto hrefs and drops javascript: ones (text survives)', () => {
    const out = sanitizeClipboardHtml(
      '<a href="https://example.com">ok</a><a href="javascript:alert(1)">bad</a>' +
        '<a href="mailto:a@b.c">mail</a>',
    )
    expect(out).toContain('href="https://example.com"')
    expect(out).toContain('href="mailto:a@b.c"')
    expect(out).not.toContain('javascript:')
    // the dead anchor degrades to its plain text
    expect(out).toContain('bad')
    expect(out).not.toMatch(/<a[^>]*>bad</)
  })

  it('keeps http(s) img, degrades others to alt text', () => {
    const out = sanitizeClipboardHtml(
      '<img src="https://cdn.example/x.png" alt="pic">' +
        '<img src="file:///C:/x.png" alt="local">' +
        '<img src="javascript:alert(1)">',
    )
    expect(out).toContain('src="https://cdn.example/x.png"')
    expect(out).toContain('alt="pic"')
    expect(out).not.toContain('file:')
    expect(out).toContain('local')
    expect(out).not.toContain('<img src="javascript')
  })

  it('strips class/id/data attributes, keeps table spans', () => {
    const out = sanitizeClipboardHtml(
      '<p class="MsoNormal" id="x" data-foo="y" dir="ltr">t</p>' +
        '<table><tr><td colspan="2">c</td></tr></table>',
    )
    expect(out).not.toContain('class=')
    expect(out).not.toContain('id=')
    expect(out).not.toContain('data-')
    expect(out).not.toContain('dir=')
    expect(out).toContain('<td colspan="2">c</td>')
  })
})

describe('degenerate inputs', () => {
  it('returns an empty string for pure junk (caller falls back to text/plain)', () => {
    expect(sanitizeClipboardHtml('<style>.x{}</style>')).toBe('')
    expect(sanitizeClipboardHtml('<!--StartFragment--><!--EndFragment-->')).toBe('')
    expect(sanitizeClipboardHtml('')).toBe('')
  })

  it('never throws on malformed markup', () => {
    expect(() => sanitizeClipboardHtml('<p>unclosed <b>bold')).not.toThrow()
  })

  it('still reports content that only consists of a line break or image', () => {
    expect(sanitizeClipboardHtml('<p><br></p>')).toContain('<br>')
    expect(sanitizeClipboardHtml('<img src="https://x/y.png">')).toContain('<img')
  })
})

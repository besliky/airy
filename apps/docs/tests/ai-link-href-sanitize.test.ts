import { describe, expect, it } from 'vitest'
import { parseInlineFragment, sanitizeLinkHref } from '../src/renderer/ai/protocol'

describe('link href sanitization', () => {
  it('keeps http, https, mailto, fragment and relative hrefs', () => {
    expect(sanitizeLinkHref('https://example.com/a?b=1')).toBe('https://example.com/a?b=1')
    expect(sanitizeLinkHref('http://example.com')).toBe('http://example.com')
    expect(sanitizeLinkHref('mailto:user@example.com')).toBe('mailto:user@example.com')
    expect(sanitizeLinkHref('#section-2')).toBe('#section-2')
    expect(sanitizeLinkHref('docs/page.html')).toBe('docs/page.html')
    expect(sanitizeLinkHref('  https://example.com  ')).toBe('https://example.com')
  })

  it('drops dangerous and non-web schemes', () => {
    expect(sanitizeLinkHref('javascript:alert(1)')).toBeNull()
    expect(sanitizeLinkHref('JAVASCRIPT:alert(1)')).toBeNull()
    expect(sanitizeLinkHref('java\nscript:alert(1)'.replace('\n', ''))).toBeNull()
    expect(sanitizeLinkHref('file:///etc/passwd')).toBeNull()
    expect(sanitizeLinkHref('data:text/html,<script>1</script>')).toBeNull()
    expect(sanitizeLinkHref('vbscript:msgbox')).toBeNull()
    expect(sanitizeLinkHref('')).toBeNull()
    expect(sanitizeLinkHref(null)).toBeNull()
    expect(sanitizeLinkHref('   ')).toBeNull()
  })

  it('persists safe hrefs into link marks of parsed fragments', () => {
    const nodes = parseInlineFragment('<a href="https://example.com">go</a>')
    expect(nodes).toEqual([
      {
        type: 'text',
        text: 'go',
        marks: [{ type: 'link', attrs: { href: 'https://example.com', rId: null } }],
      },
    ])
  })

  it('keeps the text but drops the link mark for dangerous hrefs', () => {
    const nodes = parseInlineFragment('<a href="javascript:alert(1)">go</a>')
    expect(nodes).toEqual([{ type: 'text', text: 'go' }])
    const file = parseInlineFragment('<a href="file:///etc/passwd">secret</a>')
    expect(file).toEqual([{ type: 'text', text: 'secret' }])
  })
})

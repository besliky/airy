import { describe, expect, it, vi } from 'vitest'

import { buildStandaloneHtml, sanitizeLinkHref } from '../src/renderer/export/htmlExport'

// The KaTeX stylesheet is inlined with Vite's `?inline` query; stub it so the
// test asserts the builder's own behavior rather than the bundler's.
vi.mock('katex/dist/katex.min.css?inline', () => ({ default: '' }))

function editorRoot(innerHtml: string): HTMLElement {
  const root = document.createElement('div')
  root.setAttribute('contenteditable', 'true')
  root.innerHTML = innerHtml
  return root
}

describe('sanitizeLinkHref', () => {
  it('keeps web, mail, fragment, and relative refs', () => {
    expect(sanitizeLinkHref('https://example.com/a?b=1')).toBe('https://example.com/a?b=1')
    expect(sanitizeLinkHref('http://example.com')).toBe('http://example.com')
    expect(sanitizeLinkHref('mailto:a@b.c')).toBe('mailto:a@b.c')
    expect(sanitizeLinkHref('#section')).toBe('#section')
    expect(sanitizeLinkHref('docs/page.html')).toBe('docs/page.html')
  })

  it('rejects executable and app-owned schemes plus empty values', () => {
    expect(sanitizeLinkHref('javascript:alert(1)')).toBeNull()
    expect(sanitizeLinkHref('FILE:///etc/passwd')).toBeNull()
    expect(sanitizeLinkHref('data:text/html,<b>x</b>')).toBeNull()
    expect(sanitizeLinkHref('md-asset:///doc/assets/img.png')).toBeNull()
    expect(sanitizeLinkHref(' vbscript:x ')).toBeNull()
    expect(sanitizeLinkHref('')).toBeNull()
    expect(sanitizeLinkHref(null)).toBeNull()
  })
})

describe('buildStandaloneHtml', () => {
  it('builds a standalone document: doctype, charset, title, inline styles, no base tag', async () => {
    const html = await buildStandaloneHtml(editorRoot('<h1>Hello</h1><p>World</p>'), 'Notes')
    expect(html.startsWith('<!doctype html>')).toBe(true)
    expect(html).toContain('<meta charset="utf-8">')
    expect(html).toContain('<title>Notes</title>')
    // KaTeX stylesheet plus the shared print-theme stylesheet, both inline
    expect(html.match(/<style>/g)?.length).toBeGreaterThanOrEqual(2)
    expect(html).not.toContain('<base')
    expect(html).toContain('<h1>Hello</h1>')
    expect(html).toContain('<p>World</p>')
  })

  it('escapes the document title', async () => {
    const html = await buildStandaloneHtml(editorRoot('<p>x</p>'), 'A&B <C>')
    expect(html).toContain('<title>A&amp;B &lt;C></title>')
    expect(html).not.toContain('<title>A&B <C></title>')
  })

  it('strips editor-only chrome from the clone', async () => {
    const html = await buildStandaloneHtml(
      editorRoot(
        '<p contenteditable="true">text</p>' +
          '<div class="md-codeblock-bar"><button>copy</button></div>',
      ),
      'Notes',
    )
    expect(html).not.toContain('contenteditable')
    expect(html).not.toContain('md-codeblock-bar')
    expect(html).toContain('<p>text</p>')
  })

  it('contains no live scripts and no event-handler attributes', async () => {
    const html = await buildStandaloneHtml(
      editorRoot(
        '<p onclick="steal()">t</p><script>window.__x=1</script>' +
          '<span data-raw-html="&lt;script&gt;dan ger&lt;/script&gt;">&lt;script&gt;danger&lt;/script&gt;</span>',
      ),
      'Notes',
    )
    expect(html).not.toContain('<script')
    expect(html).not.toContain('onclick=')
    // the raw-HTML chip keeps its escaped text but drops the internal carrier
    // attribute (its value holds literal markup, unescaped in attribute position)
    expect(html).not.toContain('data-raw-html')
    expect(html).toContain('&lt;script&gt;danger&lt;/script&gt;')
  })

  it('keeps safe hrefs and degrades unsafe anchors to plain text', async () => {
    const html = await buildStandaloneHtml(
      editorRoot(
        '<a href="https://example.com">web</a>' +
          '<a href="#sec">frag</a>' +
          '<a href="javascript:alert(1)">evil</a>' +
          '<a href="file:///etc/passwd">local</a>',
      ),
      'Notes',
    )
    expect(html).toContain('<a href="https://example.com">web</a>')
    expect(html).toContain('<a href="#sec">frag</a>')
    expect(html).not.toContain('javascript:')
    expect(html).not.toContain('file://')
    // the text survives the unwrap, the link does not
    expect(html).toContain('evil')
    expect(html).toContain('local')
  })

  it('inlines md-asset images as data URIs and leaves other sources untouched', async () => {
    const calls: string[] = []
    const html = await buildStandaloneHtml(
      editorRoot(
        '<img src="md-asset:///doc/assets/pic.png" alt="pic">' +
          '<img src="https://cdn.example.com/remote.png" alt="remote">' +
          '<img src="md-asset:///doc/assets/missing.png" alt="gone">',
      ),
      'Notes',
      async (src) => {
        calls.push(src)
        // the app-side loader only serves local document assets
        if (!src.startsWith('md-asset://') || src.includes('missing')) return null
        return { base64: 'QUJD', mime: 'image/png' }
      },
    )
    expect(calls).toEqual([
      'md-asset:///doc/assets/pic.png',
      'https://cdn.example.com/remote.png',
      'md-asset:///doc/assets/missing.png',
    ])
    expect(html).toContain('src="data:image/png;base64,QUJD"')
    expect(html).not.toContain('md-asset:///doc/assets/pic.png')
    expect(html).toContain('src="https://cdn.example.com/remote.png"')
    // unresolvable local image keeps its reference (alt text still shows)
    expect(html).toContain('src="md-asset:///doc/assets/missing.png"')
  })

  it('rejects loader payloads that are not images', async () => {
    const html = await buildStandaloneHtml(
      editorRoot('<img src="md-asset:///doc/assets/x.png" alt="x">'),
      'Notes',
      async () => ({ base64: 'PHNjcmlwdD4=', mime: 'text/html' }),
    )
    expect(html).not.toContain('data:')
  })
})

import katexCss from 'katex/dist/katex.min.css?inline'
import { PRINT_CSS } from './printHtml'

/**
 * Link href whitelist for the standalone export (docs BUG-406 parity): the
 * editor stores raw hrefs for fidelity and only gates them at open time via
 * safeExternalUrl — a written .html file has no such gate, so unsafe schemes
 * are filtered here. https/http/mailto, in-document fragments, and
 * scheme-less relative refs stay; everything else (javascript:, file:,
 * data:, md-asset:) is rejected.
 */
export function sanitizeLinkHref(raw: string | null): string | null {
  const href = (raw ?? '').trim()
  if (!href) return null
  if (/^(https?|mailto):/i.test(href)) return href
  if (href.startsWith('#')) return href
  // anything else with a scheme is not allowed; scheme-less values are relative refs
  if (/^[a-z][a-z0-9+.-]*:/i.test(href)) return null
  return href
}

/**
 * Display src of an image → embeddable bytes; null → the src is kept as-is.
 * Lets the caller decide which sources are inlinable (the app inlines only
 * md-asset:// images via the main process's readImage grant).
 */
export type InlineImageLoader = (src: string) => Promise<{ base64: string; mime: string } | null>

/**
 * PAR-315: standalone HTML export. The live editor DOM runs through the same
 * self-contained pipeline as the PDF/print path (printHtml's stylesheet plus
 * the inlined KaTeX CSS), with app-owned references rewritten so the file
 * opens offline in any browser:
 * - no <base> tag (the app's renderer URL would be a dead reference);
 * - md-asset:// images are inlined as data URIs via the loader;
 * - link hrefs run through the scheme whitelist — an anchor whose href is
 *   rejected degrades to plain text, exactly like the docs export;
 * - editor chrome (contenteditable, code block bars) never makes it into
 *   the clone, and scripts/event-handler attributes are stripped as defense
 *   in depth (the editor already renders raw HTML as escaped text).
 */
export async function buildStandaloneHtml(
  editorRoot: HTMLElement,
  title: string,
  loadImage?: InlineImageLoader,
): Promise<string> {
  const clone = editorRoot.cloneNode(true) as HTMLElement
  clone.removeAttribute('contenteditable')
  for (const el of clone.querySelectorAll('[contenteditable]'))
    el.removeAttribute('contenteditable')

  // editor-only chrome (language picker + copy button, mermaid failure note)
  // must not appear in the export; a rendered mermaid block exports as its diagram
  for (const bar of clone.querySelectorAll('.md-codeblock-bar, .md-mermaid-error')) bar.remove()
  for (const block of clone.querySelectorAll('[data-mermaid="rendered"] pre')) block.remove()

  for (const script of clone.querySelectorAll('script')) script.remove()
  for (const el of clone.querySelectorAll('*')) {
    for (const attr of [...el.attributes]) {
      // on*: event handlers must never attach; data-raw-html: the raw-HTML
      // chip's verbatim-markup carrier — the HTML serializer does not escape
      // `<` inside attribute values, so it would leak literal `<script>`
      // markup into the file. The chip's escaped text child stays visible.
      if (/^on/i.test(attr.name) || attr.name === 'data-raw-html') el.removeAttribute(attr.name)
    }
  }

  for (const a of [...clone.querySelectorAll('a')]) {
    if (sanitizeLinkHref(a.getAttribute('href')) === null) {
      // unwrap: the anchor's text stays, only the link is dropped
      const parent = a.parentNode
      if (parent) {
        while (a.firstChild) parent.insertBefore(a.firstChild, a)
        a.remove()
      }
    }
  }

  if (loadImage) {
    for (const img of [...clone.querySelectorAll('img')]) {
      const src = img.getAttribute('src')
      if (!src) continue
      const data = await loadImage(src).catch(() => null)
      if (!data || !/^image\//.test(data.mime)) continue
      img.setAttribute('src', `data:${data.mime};base64,${data.base64}`)
    }
  }

  const escapedTitle = title.replace(/&/g, '&amp;').replace(/</g, '&lt;')
  return [
    '<!doctype html>',
    '<html>',
    '<head>',
    '<meta charset="utf-8">',
    `<title>${escapedTitle}</title>`,
    `<style>${katexCss}</style>`,
    `<style>${PRINT_CSS}</style>`,
    '</head>',
    `<body>${clone.innerHTML}</body>`,
    '</html>',
  ].join('\n')
}

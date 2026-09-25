/**
 * UX-1705: sanitizer for clipboard HTML pasted from rich sources (Word, web
 * pages). Runs BEFORE the schema parse so Word's private markup never reaches
 * the RawHtml catch-all as junk chips.
 *
 * Sanitization policy (a deliberate subset of what the rawHtml node tolerates
 * for FILES — pasted content is foreign, so it is stricter):
 *
 * 1. Comments are removed entirely — Word's `<!--[if gte mso 9]>` data islands
 *    and the `<!--StartFragment-->` / `<!--EndFragment-->` markers die here.
 * 2. Dangerous / non-content elements are removed WITH their subtree: script,
 *    style, iframe, object, svg/math islands, form controls, media.
 * 3. Everything else not on the allow-list is UNWRAPPED (children kept,
 *    element dropped): Word's `<span style='font-family:...'>` wrappers,
 *    `<font>`, `<o:p>` and every namespace-prefixed (`st1:*`, `v:*`, `w:*`)
 *    office tag. This is the "inline styles are CUT, not simplified" choice —
 *    colors/fonts do not survive, semantics (b/i/u/h/p/li/table) do.
 * 4. Allowed elements keep NO attributes except a tiny semantic set: `a`
 *    href/title, `img` src/alt/title, `td`/`th` colspan/rowspan, `col(group)`
 *    span, `abbr` title. Every `on*` handler, `class`, `id`, `style` and
 *    `data-*` attribute is stripped by not being on the list. `href`/`src`
 *    must be http(s) or mailto — `javascript:` URLs drop the attribute (the
 *    anchor itself degrades to plain text), and an `img` without a safe src
 *    degrades to its alt text.
 *
 * The output is an HTML fragment string (body innerHTML) that the editor's own
 * parse rules then turn into schema nodes: dedicated rules win (strong/em/
 * headings/tables/...), and anything the schema has no node for (mark, kbd,
 * details) still lands as a rawHtml chip — the same policy as typing it.
 */

const DROP_WITH_CONTENT = new Set([
  'script',
  'style',
  'link',
  'meta',
  'base',
  'title',
  'template',
  'noscript',
  'iframe',
  'frame',
  'frameset',
  'object',
  'embed',
  'applet',
  'svg',
  'math',
  'form',
  'input',
  'button',
  'select',
  'option',
  'optgroup',
  'textarea',
  'audio',
  'video',
  'source',
  'track',
  'canvas',
  'map',
  'area',
  'dialog',
  'slot',
])

const ALLOWED = new Set([
  // block structure
  'p',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'blockquote',
  'pre',
  'br',
  'hr',
  'ul',
  'ol',
  'li',
  'dl',
  'dt',
  'dd',
  // tables
  'table',
  'thead',
  'tbody',
  'tfoot',
  'tr',
  'td',
  'th',
  'caption',
  'col',
  'colgroup',
  // inline formatting
  'b',
  'strong',
  'i',
  'em',
  'u',
  's',
  'strike',
  'del',
  'ins',
  'code',
  'kbd',
  'mark',
  'sub',
  'sup',
  'abbr',
  'q',
  'cite',
  'dfn',
  'var',
  'samp',
  'small',
  'big',
  'tt',
  'wbr',
  // embedded semantics
  'a',
  'img',
  'details',
  'summary',
])

const KEEP_ATTRS: Record<string, ReadonlySet<string>> = {
  a: new Set(['href', 'title']),
  img: new Set(['src', 'alt', 'title']),
  td: new Set(['colspan', 'rowspan']),
  th: new Set(['colspan', 'rowspan']),
  col: new Set(['span']),
  colgroup: new Set(['span']),
  abbr: new Set(['title']),
}

const SAFE_URL = /^(https?:|mailto:)/i

/**
 * Strip a clipboard HTML payload down to the allow-listed fragment.
 * Returns '' when nothing content-worthy survives (the caller then falls back
 * to the text/plain flavor).
 */
export function sanitizeClipboardHtml(html: string): string {
  if (!html || !html.trim()) return ''
  let doc: Document
  try {
    doc = new DOMParser().parseFromString(html, 'text/html')
  } catch {
    return ''
  }
  const body = doc.body
  if (!body) return ''

  // (1) comments: conditional mso blocks, StartFragment/EndFragment markers
  const walker = doc.createTreeWalker(body, NodeFilter.SHOW_COMMENT)
  const comments: Comment[] = []
  for (let node = walker.nextNode(); node; node = walker.nextNode()) comments.push(node as Comment)
  for (const comment of comments) comment.remove()

  // (2)/(3)/(4): post-order — children are cleaned before their parent is
  // judged, so unwrapping only ever re-parents already-clean content
  const clean = (element: Element): void => {
    for (const child of Array.from(element.children)) clean(child)
    const tag = element.tagName.toLowerCase()
    if (DROP_WITH_CONTENT.has(tag)) {
      element.remove()
      return
    }
    if (!ALLOWED.has(tag)) {
      element.replaceWith(...element.childNodes)
      return
    }
    const keep = KEEP_ATTRS[tag]
    for (const attr of Array.from(element.attributes)) {
      const name = attr.name.toLowerCase()
      if (
        !keep?.has(name) ||
        ((name === 'href' || name === 'src') && !SAFE_URL.test(attr.value.trim()))
      ) {
        element.removeAttribute(attr.name)
      }
    }
  }
  for (const child of Array.from(body.children)) clean(child)

  // (4) degrade anchors without a safe href and images without a safe src:
  // dead `<a>` wrappers unwrap to their text, `<img>` collapses to its alt
  for (const anchor of Array.from(body.querySelectorAll('a'))) {
    if (!anchor.getAttribute('href')) anchor.replaceWith(...anchor.childNodes)
  }
  for (const image of Array.from(body.querySelectorAll('img'))) {
    if (!image.getAttribute('src')) {
      image.replaceWith(doc.createTextNode(image.getAttribute('alt') ?? ''))
    }
  }

  // a fragment of pure whitespace is no content — the caller falls back to plain
  if (!body.textContent?.trim() && !body.querySelector('img,br,hr')) return ''
  return body.innerHTML
}

import { Node, mergeAttributes } from '@tiptap/core'

/**
 * Verbatim preservation of raw HTML the schema has no dedicated node or mark
 * for (BUG-1685: opening a README with `<div>`/`<script>`/`<style>`/`<kbd>`/
 * `<mark>`/`<details>` and saving silently stripped every element).
 *
 * How it works inside `@tiptap/markdown`: unknown HTML reaches the schema via
 * `generateJSON` (a DOM parse against the full extension list). Every dedicated
 * parseDOM rule — strong, em, br, img, a[href], headings, tables, math — is
 * tried before the catch-all rule below (priority 10 < the default 50), so
 * schema-known markup keeps its normal model representation and the legitimate
 * `<br>`→hard-break / `<img>`→image conversions are untouched. Anything no
 * dedicated rule claims is captured by this node with the element's
 * `outerHTML` stored verbatim in the `html` attribute, and `renderMarkdown`
 * writes that attribute back unchanged — so the markup survives save, reopen
 * and the segmented hydration path (which parses through the same manager).
 *
 * Rendering uses the raw markup as ESCAPED TEXT (a plain text child), never
 * `innerHTML` — scripts stored in the document must not execute in the editor
 * (audit: `window.__scriptRan === false` is a hard requirement), and event-
 * handler attributes must not get a chance to attach.
 */
export const RawHtml = Node.create({
  name: 'rawHtml',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  // no marks: the chip stands for source markup, decorating it would be lost
  // on save anyway
  marks: '',

  addAttributes() {
    return {
      html: {
        default: '',
        // a re-parse of the editor's own DOM (internal clipboard round-trip)
        // finds the original markup in the data attribute; a foreign element
        // is captured with its full outer HTML
        parseHTML: (element) => element.getAttribute('data-raw-html') ?? element.outerHTML,
        renderHTML: (attributes) => ({ 'data-raw-html': String(attributes.html ?? '') }),
      },
    }
  },

  parseHTML() {
    // catch-all: every element matches, but only after all dedicated rules
    // declined (lower priority) — and only for elements the schema would
    // otherwise DROP, which is exactly the loss this node exists to prevent
    return [{ tag: '*', priority: 10 }]
  },

  renderHTML({ node }) {
    // text child → the browser escapes it; the raw markup is displayed as
    // source, matching the "verbatim preservation" contract
    return [
      'span',
      mergeAttributes(this.options.HTMLAttributes, { class: 'raw-html-chip' }),
      String(node.attrs.html ?? ''),
    ]
  },

  // serializer: write the stored markup back byte-for-byte (block-shaped
  // chunks land in their own paragraph and parse back as raw HTML blocks;
  // inline ones merge back exactly like the source tokens did)
  renderMarkdown: (node) => String(node.attrs?.html ?? ''),
})

/** XML escaping utilities (same as docx-engine's, used for patch generation). */

/**
 * A fast-xml-parser element node: attributes are '@_'-prefixed, text is '#text',
 * child elements are nested nodes (a single child collapses to an object instead
 * of an array). Parse trees are inherently untyped, so readers probe them through
 * the narrowing helpers below instead of `any`.
 */
export type XmlNode = Record<string, unknown>

/** View an unknown parse-tree value as an element node; non-objects read as empty. */
export function asXmlNode(v: unknown): XmlNode {
  return typeof v === 'object' && v !== null ? (v as XmlNode) : {}
}

/** Normalize fast-xml-parser's single-child collapse: always get an array of element nodes. */
export function xmlArray(v: unknown): XmlNode[] {
  if (Array.isArray(v)) return v.map(asXmlNode)
  return v ? [asXmlNode(v)] : []
}

/** XML 1.0 forbids C0 controls (minus tab/LF/CR), U+FFFE/FFFF and lone
    surrogates even when escaped — one such byte makes the whole part
    unparseable and PowerPoint offers repair. VT/FF (common in PDF-extracted
    text) degrade to a space; the rest have no textual meaning and drop. */
// eslint-disable-next-line no-control-regex -- the forbidden chars are the subject
const XML_INVALID_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|\p{Cs}/gu

function sanitizeXmlChars(text: string): string {
  XML_INVALID_CHARS.lastIndex = 0
  if (!XML_INVALID_CHARS.test(text)) return text
  XML_INVALID_CHARS.lastIndex = 0
  return text.replace(XML_INVALID_CHARS, (ch) => (ch === '\u000B' || ch === '\u000C' ? ' ' : ''))
}

export function escapeXmlText(text: string): string {
  return sanitizeXmlChars(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function escapeXmlAttr(text: string): string {
  return escapeXmlText(text).replace(/"/g, '&quot;')
}

/**
 * a16:creationId extLst for a newborn <p:cNvPr> — durable identity from birth
 * (design step 0): the GUID is written into the file bytes, so ids survive
 * save→reopen, reparse, group/ungroup, and editors that renumber cNvPr ids.
 * The a16 namespace is declared inline so the fragment is valid standalone.
 */
export function creationIdExtXml(): string {
  return (
    '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{${globalThis.crypto.randomUUID().toUpperCase()}}"/>` +
    '</a:ext>'
  )
}

export function creationIdXml(): string {
  return `<a:extLst>${creationIdExtXml()}</a:extLst>`
}

/**
 * Ext uri marking a chart frame as created/claimed by this app: the chart's
 * <p:cNvPr> carries this marker inside its <a:extLst>, so the chart is offered
 * for in-app editing on reopen. It deliberately does NOT use cNvPr@descr —
 * that attribute is the user's alt-text description slot (PAR-304); files from
 * older versions with descr="aislides-chart" are still recognized on read.
 */
export const APP_CHART_MARKER_URI = '{7B1D4C62-95AE-4E68-9D3F-52C08A41E7B4}'

/**
 * The marker ext fragment. The aislides namespace is declared inline so the
 * fragment stays valid wherever it is spliced (like creationIdExtXml).
 */
export function appChartMarkerExtXml(): string {
  return (
    `<a:ext uri="${APP_CHART_MARKER_URI}">` +
    '<aislides:chart xmlns:aislides="http://schemas.airy-office.app/drawing/2026/main"/>' +
    '</a:ext>'
  )
}

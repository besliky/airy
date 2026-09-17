import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

/** ProseMirror positions of a top-level child index range */
export function blockRangePositions(
  editor: Editor,
  startIndex: number,
  endIndex: number,
): { from: number; to: number } {
  const doc = editor.state.doc
  let from = 0
  let to = 0
  let index = 0
  doc.forEach((node, offset) => {
    if (index === startIndex) from = offset
    if (index === endIndex) to = offset + node.nodeSize
    index++
  })
  return { from, to }
}

// ---- tracked deletions (pending revisions are not current content) ----

const hasDelMark = (node: ProseMirrorNode) => node.marks.some((m) => m.type.name === 'del')

// ---- link href policy (shared by the HTML and ops entry points) ----

/**
 * Whitelist for hrefs entering link marks: http(s), mailto, in-document
 * fragments and plain relative references. Everything else — javascript:,
 * file:, data:, vbscript: … — is dropped so a prompt-injected answer cannot
 * persist a dangerous scheme into the document (opening is separately gated,
 * but the stored attribute itself must stay benign). Returns null to mean
 * "keep the text, drop the link".
 *
 * Residuals (accepted): this policy covers the AI/bridge/ops entry points
 * only. Three user-side or file-load paths persist raw hrefs by design —
 * pasted-HTML link marks (LinkMark.parseHTML in editor/marks.ts), the
 * Insert-Link dialog (components/ribbon-insert-tab.tsx), and docx file
 * loading (editor/convert.ts, for fidelity with what the file contains).
 * There the user, not a model, supplied the value; every link OPENING path
 * still routes through safeExternalUrl, which rejects non-http(s)/mailto
 * schemes, so a stored javascript: href can never execute (see SECURITY.md).
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

/** block text as it reads once pending tracked deletions are applied (textContent minus del runs) */
export function liveText(node: ProseMirrorNode): string {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return ''
  let out = ''
  const walk = (child: ProseMirrorNode): void => {
    if (hasDelMark(child)) return
    if (child.isText) out += child.text ?? ''
    else if (child.isLeaf) out += child.type.spec.leafText?.(child) ?? ''
    else child.forEach(walk)
  }
  node.forEach(walk)
  return out
}

/**
 * The whole block is a pending deletion revision (struck through in the editor).
 * Checked structurally, not by text length: atom leaves (inline formulas, ruby)
 * carry no textContent, so a live formula must still count as live content.
 */
export function isTrackedDeleted(node: ProseMirrorNode): boolean {
  if ((node.attrs?.blockRevision as { kind?: string } | null)?.kind === 'del') return true
  let hasContent = false
  let hasLive = false
  node.descendants((child) => {
    if (child.isText || (child.isInline && child.isLeaf)) {
      hasContent = true
      if (!hasDelMark(child)) hasLive = true
    }
  })
  return hasContent && !hasLive
}

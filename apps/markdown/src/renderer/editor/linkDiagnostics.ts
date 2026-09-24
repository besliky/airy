import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { showToast } from '@airy-office/ui/toast-bus'
import { t } from '../i18n/locale'

/**
 * Honest, view-only diagnostics for links/anchors the markdown parser cannot
 * resolve (UX-1701; P4 tail of the 2026-09-24 markdown audit). Nothing here
 * changes the document model, the serializer or the saved file — it only adds
 * decorations:
 *
 * - `[b]` used with a circular/undefined reference definition (`[b]: [a]`)
 *   parses to href `[a]` (the literal label, not a URL) — marked as an
 *   unresolved reference instead of silently looking like a working link.
 * - `[text](#anchor)` where no heading (or raw-HTML `<a id/name>`) produces
 *   that anchor — marked as a dead anchor, unobtrusive underline.
 * - `[empty]()` — href is the empty string; marked as an empty link.
 * - `[^1]` footnotes are not supported (no footnote extension); they stay
 *   literal text. They get a subtle hint decoration, and the first appearance
 *   in a document raises one toast so the limitation is not silent.
 *
 * Anchor targets follow the GitHub heading-slug convention (the de-facto
 * standard for markdown files), plus explicit ids in raw-HTML nodes kept by
 * the RawHtml extension.
 */

export const linkDiagnosticsPluginKey = new PluginKey<{
  set: DecorationSet
  footnoteFound: boolean
}>('mdLinkDiagnostics')

/** Upper bound on painted diagnostics, same rationale as MAX_PAINTED_HITS in findTarget */
const MAX_PAINTED = 500

/** `[^label]` — footnote syntax the schema has no node or mark for */
const FOOTNOTE_RE = /\[\^[^\]\s]+\]/g

/** href the parser produced for a reference link whose definition resolves to another label (`[b]: [a]`) */
const REFERENCE_LABEL_RE = /^\[[^\]\s]*\]$/

/** GitHub heading anchor: lowercase, drop punctuation, spaces become hyphens */
export function headingSlug(text: string): string {
  return text
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\p{M}\s-]/gu, '')
    .replace(/\s/g, '-')
}

/** Collect every in-document anchor target: heading slugs plus raw-HTML id/name attributes */
function collectAnchors(doc: ProseMirrorNode): Set<string> {
  const anchors = new Set<string>()
  doc.descendants((node) => {
    if (node.type.name === 'heading') {
      anchors.add(headingSlug(node.textContent))
    } else if (node.type.name === 'rawHtml') {
      // `<a id="x">`/`<a name="x">` kept verbatim by RawHtml still anchors #x
      const html = String(node.attrs.html ?? '')
      for (const match of html.matchAll(/\b(?:id|name)\s*=\s*["']([^"']+)["']/gi)) {
        anchors.add(match[1])
      }
    }
    return true
  })
  return anchors
}

export type LinkIssue = 'empty' | 'unresolvedRef' | 'deadAnchor'

/** Classify an unresolvable link href; null when the link is fine */
export function linkIssue(href: string, anchors: ReadonlySet<string>): LinkIssue | null {
  if (href === '') return 'empty'
  if (REFERENCE_LABEL_RE.test(href)) return 'unresolvedRef'
  if (href.startsWith('#')) {
    let anchor = href.slice(1)
    try {
      anchor = decodeURIComponent(anchor)
    } catch {
      // malformed percent-encoding — compare the raw form
    }
    if (!anchors.has(anchor)) return 'deadAnchor'
  }
  return null
}

const ISSUE_CLASS: Record<LinkIssue, string> = {
  empty: 'md-link-empty',
  unresolvedRef: 'md-link-unresolved-ref',
  deadAnchor: 'md-link-dead-anchor',
}

const ISSUE_TITLE_KEY: Record<LinkIssue, 'linkEmpty' | 'linkUnresolvedRef' | 'linkDeadAnchor'> = {
  empty: 'linkEmpty',
  unresolvedRef: 'linkUnresolvedRef',
  deadAnchor: 'linkDeadAnchor',
}

interface ScanResult {
  set: DecorationSet
  footnoteFound: boolean
}

function scanDoc(doc: ProseMirrorNode): ScanResult {
  const anchors = collectAnchors(doc)
  const decorations: Decoration[] = []
  let footnoteFound = false
  doc.descendants((node, pos) => {
    if (decorations.length > MAX_PAINTED) return false // cap reached — stop walking
    const text = node.isText ? node.text : null
    if (!text) return true
    const linkMark = node.marks.find((mark) => mark.type.name === 'link')
    if (linkMark) {
      const issue = linkIssue(String(linkMark.attrs.href ?? ''), anchors)
      if (issue) {
        decorations.push(
          Decoration.inline(pos, pos + node.nodeSize, {
            class: ISSUE_CLASS[issue],
            title: t(ISSUE_TITLE_KEY[issue]),
          }),
        )
      }
    }
    if (text.includes('[^')) {
      for (const match of text.matchAll(FOOTNOTE_RE)) {
        footnoteFound = true
        const start = pos + (match.index ?? 0)
        decorations.push(
          Decoration.inline(start, start + match[0].length, {
            class: 'md-footnote-hint',
            title: t('footnoteUnsupported'),
          }),
        )
      }
    }
    return true
  })
  return {
    set: DecorationSet.create(doc, decorations.slice(0, MAX_PAINTED)),
    footnoteFound,
  }
}

/** One toast per document: the first time a literal footnote shows up, say so */
function notifyFootnoteOnce(editor: Editor, storage: { footnoteToasted: boolean }): void {
  if (storage.footnoteToasted) return
  if (linkDiagnosticsPluginKey.getState(editor.state)?.footnoteFound) {
    storage.footnoteToasted = true
    showToast(t('footnoteUnsupported'), 'error')
  }
}

export const LinkDiagnostics = Extension.create({
  name: 'linkDiagnostics',

  addStorage() {
    return { footnoteToasted: false }
  },

  onCreate() {
    notifyFootnoteOnce(this.editor, this.storage)
  },

  onUpdate() {
    notifyFootnoteOnce(this.editor, this.storage)
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: linkDiagnosticsPluginKey,
        state: {
          init: (_, state) => scanDoc(state.doc),
          apply(tr, old, _oldState, newState) {
            // selection-only transactions never change diagnostics
            if (!tr.docChanged) return old
            return scanDoc(newState.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)?.set ?? DecorationSet.empty
          },
        },
      }),
    ]
  },
})

import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { sanitizeClipboardHtml } from './pasteSanitize'

/**
 * UX-1705: "paste from Word with formatting". A clipboard carrying the
 * `text/html` flavor is sanitized ({@link sanitizeClipboardHtml}) and inserted
 * through the schema's own parse rules — dedicated rules keep winning (b/i/u
 * become marks, p/h/li/table become nodes) and what the schema has no node for
 * survives as a rawHtml chip, exactly the policy the rawHtml node defines.
 * A plain paste (no html flavor) behaves exactly as before.
 *
 * Mod+Shift+V (and the ribbon command) force the text/plain flavor: the
 * keydown only RECORDS the gesture — the browser then fires the paste event
 * and the handler below consumes the flag, so plain insertion goes through the
 * same synchronous pipeline without any async clipboard access.
 */

/** how long after the Mod+Shift+V keydown a paste event still counts as plain */
const PLAIN_PASTE_WINDOW_MS = 1500

let plainPasteUntil = 0

/** keydown side: recognize Mod+Shift+V and arm the next paste as plain */
export function notePlainPasteGesture(event: KeyboardEvent): void {
  if (!(event.ctrlKey || event.metaKey) || !event.shiftKey || event.altKey) return
  if (event.key.toLowerCase() !== 'v') return
  plainPasteUntil = Date.now() + PLAIN_PASTE_WINDOW_MS
}

/** paste side: is the armed gesture active? Always disarms after the read. */
export function consumePlainPasteGesture(): boolean {
  const active = Date.now() < plainPasteUntil
  plainPasteUntil = 0
  return active
}

/**
 * Insert text verbatim: no parsing, no marks, newlines become hard breaks.
 * Shared by the Mod+Shift+V path, the sanitizer fallback and the ribbon
 * "paste as plain text" command.
 */
export function insertPlainText(editor: Editor, text: string): void {
  const normalized = text.replace(/\r\n?/g, '\n')
  if (!normalized) return
  const lines = normalized.split('\n')
  const content = lines.flatMap((line, i) => {
    const parts: Array<Record<string, string>> = []
    if (line) parts.push({ type: 'text', text: line })
    if (i < lines.length - 1) parts.push({ type: 'hardBreak' })
    return parts
  })
  if (!content.length) return
  editor.chain().focus().insertContent(content).run()
}

export const RichPaste = Extension.create({
  name: 'richPaste',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key: new PluginKey('richPaste'),
        props: {
          handleKeyDown(_view, event) {
            notePlainPasteGesture(event)
            return false
          },
          handlePaste(_view, event) {
            const data = event.clipboardData
            if (!data) return false
            const plain = data.getData('text/plain') ?? ''

            // Mod+Shift+V: the user explicitly asked for the plain flavor
            if (consumePlainPasteGesture()) {
              if (!plain) return false
              insertPlainText(editor, plain)
              return true
            }

            const html = data.getData('text/html')
            // no html flavor — the historic plain-text path, untouched
            if (!html) return false
            // ProseMirror's own serialized slice (internal copy): let the
            // default exact-slice restore run, or image nodes and rawHtml
            // chips would degrade through the sanitizer
            if (html.includes('data-pm-slice')) return false
            // inside a code block markup must land as source text — the
            // default handler already does exactly that
            if (editor.isActive('codeBlock')) return false

            const sanitized = sanitizeClipboardHtml(html)
            if (!sanitized) {
              // nothing content-worthy survived (pure mso junk) — plain fallback
              if (!plain) return false
              insertPlainText(editor, plain)
              return true
            }
            editor.chain().focus().insertContent(sanitized).run()
            return true
          },
        },
      }),
    ]
  },
})

import type { EditorView } from '@codemirror/view'
import { sanitizeClipboardHtml } from './pasteSanitize'

/**
 * UX-1705: paste routing for the HTML source editor. A clipboard with the
 * `text/html` flavor (a Word fragment, a selection copied from a rendered
 * page) is sanitized with the same allow-list policy as the markdown app and
 * inserted as markup text — previously the flavor was ignored and only
 * text/plain reached the buffer, so a Word paste lost every tag. A clipboard
 * without the flavor (including CodeMirror's own plain-text copies) behaves
 * exactly as before.
 *
 * Mod+Shift+V (and the ribbon command) force the text/plain flavor. The
 * keydown only RECORDS the gesture — the browser then fires the paste event
 * and {@link sourcePaste} consumes the flag, so the plain path needs no async
 * clipboard access; CodeMirror's own default paste is already plain, so an
 * armed gesture with an empty text flavor just falls through.
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

/** replace the selection with `text` as one undoable paste-shaped step */
export function insertTextAtSelection(view: EditorView, text: string): void {
  const { from, to } = view.state.selection.main
  view.dispatch({
    changes: { from, to, insert: text },
    selection: { anchor: from + text.length },
    userEvent: 'input.paste',
  })
}

/**
 * `paste` DOM handler (CodeMirror domEventHandlers shape). Returns true when
 * the event was handled — CodeMirror then suppresses its default insertion.
 */
export function sourcePaste(view: EditorView, event: ClipboardEvent): boolean {
  const data = event.clipboardData
  if (!data) return false
  const plain = data.getData('text/plain') ?? ''

  // Mod+Shift+V: the user explicitly asked for the plain flavor; CodeMirror's
  // default paste is already text/plain, so only an empty flavor falls through
  if (consumePlainPasteGesture()) {
    if (!plain) return false
    insertTextAtSelection(view, plain)
    return true
  }

  const html = data.getData('text/html')
  if (!html) return false
  // nothing content-worthy survived the sanitizer (pure mso junk) — plain fallback
  const insert = sanitizeClipboardHtml(html) || plain
  if (!insert) return false
  insertTextAtSelection(view, insert)
  return true
}

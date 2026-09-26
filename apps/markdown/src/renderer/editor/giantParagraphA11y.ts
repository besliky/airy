import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import type { Node as PMNode } from '@tiptap/pm/model'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { t } from '../i18n/locale'
import { GIANT_PARAGRAPH_CHARS } from './giantTextChunking'

/**
 * Accessibility policy for giant paragraphs (PERF-1700b follow-up).
 *
 * The shell forces Electron accessibility support on unconditionally
 * (`app.setAccessibilitySupportEnabled(true)`), so every renderer pays the
 * Chromium accessibility lifecycle on every change. Tracing a keystroke on a
 * multi-megabyte single-line paragraph shows the entire user-visible stall —
 * ~2.0-2.3s, previously misattributed to the software compositor — is
 * `LocalFrameView::RunAccessibilitySteps` inside the BeginMainFrame commit:
 * the AX tree update recomputes for the whole document, scales linearly with
 * the giant paragraph's text volume, and runs even when the edit lands in a
 * different paragraph. Because it happens in the commit/lifecycle phase it
 * never appears as a longtask, which is why it evaded per-keystroke
 * long-task attribution.
 *
 * Policy: paragraphs at or above the shared giant-paragraph threshold
 * (GIANT_PARAGRAPH_CHARS, 100k chars — the same population the DOM text
 * chunking already targets) are hidden from the accessibility tree with
 * `aria-hidden="true"` (a PM node decoration, so the attribute survives
 * prosemirror-view's attribute syncing). Measured live on the 4MB one-liner
 * (xvfb, software raster, N>=3): keystroke keydown->presented-frame drops
 * from ~2.15-2.36s to ~0.15-0.23s, undo x5 from ~2.13-2.32s to ~0.08-0.13s.
 *
 * A11y tradeoff and fallbacks, deliberate and documented: assistive
 * technology cannot browse or read these paragraphs. That content is
 * pathological for screen readers anyway (a 4MB paragraph is ~42k lines of
 * continuous text), and the cost of keeping it exposed is a ~0.3s (1MB) to
 * ~2s (4MB) tax on EVERY keystroke in the document for ALL users. As the
 * replacement, each hidden paragraph gets a visually-hidden summary widget
 * in front of it ("very long paragraph, N characters") so assistive
 * technology still sees the block's existence and size; caret navigation,
 * selection, find, copy/paste and IME are DOM-level and unaffected. Editing
 * inside a hidden giant paragraph is not announced by screen readers — split
 * such content into paragraphs to regain full accessibility.
 *
 * The decorations are view-only: nothing here enters the document model, so
 * serialization (getMarkdown/getJSON) is unaffected.
 */

interface GiantParaA11yState {
  decos: DecorationSet
}

function isGiantParagraph(node: PMNode): boolean {
  return (
    node.type.name === 'paragraph' && node.isTextblock && node.content.size >= GIANT_PARAGRAPH_CHARS
  )
}

/**
 * Builds, for every giant paragraph: the `aria-hidden` node decoration plus a
 * visually-hidden summary widget in front of it. The widget's spec key pins
 * the character count, so the summary DOM is reused across transactions and
 * rebuilt only when the count actually changes.
 */
function buildDecos(doc: PMNode): DecorationSet {
  const decos: Decoration[] = []
  doc.descendants((node, pos) => {
    if (!isGiantParagraph(node)) return
    decos.push(Decoration.node(pos, pos + node.nodeSize, { 'aria-hidden': 'true' }))
    const chars = node.content.size
    const summaryDom = (): HTMLElement => {
      const span = document.createElement('span')
      span.className = 'md-giant-para-summary'
      span.textContent = t('giantParaA11ySummary', { chars: chars.toLocaleString('en-US') })
      return span
    }
    decos.push(
      Decoration.widget(pos, summaryDom, {
        key: `md-giant-para-ax-${chars}`,
        side: -1,
        raw: true,
      }),
    )
  })
  if (decos.length === 0) return DecorationSet.empty
  return DecorationSet.create(doc, decos)
}

function applyTransaction(tr: Transaction, old: GiantParaA11yState): GiantParaA11yState {
  if (!tr.docChanged) return old
  return { decos: buildDecos(tr.doc) }
}

export const giantParagraphA11yKey = new PluginKey<GiantParaA11yState>('mdGiantParaA11y')

export const GiantParagraphA11y = Extension.create({
  name: 'giantParagraphA11y',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: giantParagraphA11yKey,
        state: {
          init: (_, state) => ({ decos: buildDecos(state.doc) }),
          apply: applyTransaction,
        },
        props: {
          decorations(state) {
            return giantParagraphA11yKey.getState(state)?.decos ?? DecorationSet.empty
          },
        },
      }),
    ]
  },
})

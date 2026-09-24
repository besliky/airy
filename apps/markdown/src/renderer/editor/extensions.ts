import type { AnyExtension } from '@tiptap/core'
import { Extension } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { Markdown } from '@tiptap/markdown'
import { TableKit } from '@tiptap/extension-table'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { CodeBlock } from '@tiptap/extension-code-block'
import { ReactNodeViewRenderer } from '@tiptap/react'
import { Placeholder } from '@tiptap/extensions'
import { CodeBlockView } from './CodeBlockView'
import { LocalImage } from './localImage'
import { RawHtml } from './rawHtml'
import { BlockDragHandle } from './blockDragHandle'
import { BlockKeymap } from './blockKeymap'
import { AiHighlight } from './aiHighlight'
import { AiQueueAnchors } from './aiQueueAnchors'
import { InactiveSelection } from './inactiveSelection'
import { SearchHighlight } from './searchHighlight'
import { LinkDiagnostics } from './linkDiagnostics'
import { GiantTextChunking } from './giantTextChunking'
import { buildMathExtensions } from './math'
import { SlashCommand } from './slashCommand'
import type { SlashController, SlashItem } from './slashCommand'
import { liftIndentedCodeAfterLists, stripBlankLinePadding } from '../markdown/parseContext'
import { t } from '../i18n/locale'

export interface BuildExtensionsOptions {
  slashController: SlashController
  slashItems: () => SlashItem[]
}

/**
 * App-level parse/serialize context for the markdown manager (BUG-1703). The
 * manager has no pre/post hooks, so the two pure transforms are wrapped
 * around its methods once it exists — this extension must stay AFTER the
 * `Markdown` extension in the list, because the manager is created in the
 * Markdown extension's own onBeforeCreate.
 */
const MarkdownParseContext = Extension.create({
  name: 'markdownParseContext',
  onBeforeCreate() {
    const manager = this.editor.markdown
    if (!manager) return
    const parse = manager.parse.bind(manager)
    manager.parse = (markdown: string) => parse(liftIndentedCodeAfterLists(markdown))
    const serialize = manager.serialize.bind(manager)
    manager.serialize = (doc) => stripBlankLinePadding(serialize(doc))
  },
})

export function buildExtensions(options: BuildExtensionsOptions): AnyExtension[] {
  return [
    StarterKit.configure({
      // LocalImage replaces the plain image; links open externally via main-process guard
      link: { openOnClick: false },
      // replaced by the NodeView-enhanced variant below (language picker + copy)
      codeBlock: false,
      // underline would serialize as `++text++` — not part of GFM
      underline: false,
    }),
    CodeBlock.extend({
      addNodeView() {
        return ReactNodeViewRenderer(CodeBlockView)
      },
    }),
    // 4-space nesting: the default 2 spaces is below the content column of
    // ordered items ("1. " = 3), so strict CommonMark parsers (GitHub) would
    // flatten sub-lists in the saved file. 4 is safe for every marker width.
    Markdown.configure({ indentation: { style: 'space', size: 4 } }),
    // BUG-1703 parse/serialize context fixes; must follow `Markdown` (above)
    MarkdownParseContext,
    // column widths are not expressible in GFM tables — no resizable columns;
    // the wrapper div gives wide tables a horizontal scrollbar
    TableKit.configure({ table: { resizable: false, renderWrapper: true } }),
    TaskList,
    TaskItem.configure({ nested: true }),
    // KaTeX-rendered $...$ / $$...$$ formulas (issue #100)
    ...buildMathExtensions(),
    LocalImage,
    // verbatim raw-HTML preservation (BUG-1685) — must come after the schema
    // extensions it falls back from; see rawHtml.ts
    RawHtml,
    BlockDragHandle,
    BlockKeymap,
    AiHighlight,
    AiQueueAnchors,
    InactiveSelection,
    SearchHighlight,
    // decorations for unresolvable links/anchors + the footnote hint (UX-1701);
    // view-only, never touches model or serialization
    LinkDiagnostics,
    // DOM text-node chunking for giant single paragraphs (PERF-1700); view-only
    GiantTextChunking,
    Placeholder.configure({ placeholder: () => t('placeholder') }),
    SlashCommand.configure({
      controller: options.slashController,
      items: options.slashItems,
    }),
  ]
}

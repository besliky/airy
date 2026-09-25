/**
 * F9 inline field-cache collection (PAGE / NUMPAGES / DATE / FILENAME / REF ...).
 *
 * Extracted from the App-level updateFields callback so the resolution rules are
 * unit-testable. Position rule (BUG-1709): a body PAGE field resolves from the
 * live page slicing — the page its block actually sits on — not from the
 * viewport's current page. F9 used to stamp every body PAGE field with the
 * status-bar page, so a document opened at the top silently cached "1" into
 * every PAGE result in the body.
 */
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { FieldCacheJob } from './revisions'

/** uppercased first word of a field instruction (the keyword: PAGE, REF, ...) */
export function fieldKeyword(instr: string): string {
  return instr.trim().split(/\s+/)[0]?.toUpperCase() ?? ''
}

/**
 * Position of the top-level block containing `pos`. The pagination slicer maps
 * top-level block elements to page tops, so a nested position (a field run
 * inside a paragraph, table cell, ...) is lifted to its depth-1 ancestor.
 * Returns null for depth-0 positions (there is no enclosing block).
 */
export function topLevelBlockPos(doc: PmNode, pos: number): number | null {
  const $pos = doc.resolve(pos)
  return $pos.depth >= 1 ? $pos.before(1) : null
}

export interface InlineFieldResolvers {
  /** cached value for position-independent fields (NUMPAGES / DATE / FILENAME...) */
  fieldValue: (instr: string) => string
  /** lazily built displayed-page lookup: PAGE fields and \p references are the
   * only consumers and measuring the canvas is not free, so the factory runs at
   * most once per update and only when such a field actually exists */
  pageOf: () => ((pos: number) => number | null) | null
  /** cached REF display text (null keeps the current cache). The page lookup is
   * passed for \p references, null for the others */
  refCache: (instr: string, pageOf: ((pos: number) => number | null) | null) => string | null
}

/** one FieldCacheJob per stale inline field result (instrField / refField marks) */
export function collectInlineFieldJobs(
  editor: Editor,
  resolvers: InlineFieldResolvers,
): FieldCacheJob[] {
  const jobs: FieldCacheJob[] = []
  // memoized even when the factory returns null: measuring must not repeat per field
  let pageLookup: ((pos: number) => number | null) | null | undefined
  const pageOf = (): ((pos: number) => number | null) | null => {
    if (pageLookup === undefined) pageLookup = resolvers.pageOf()
    return pageLookup
  }
  editor.state.doc.descendants((node, pos) => {
    if (!node.isText) return
    const mark = node.marks.find((m) => m.type.name === 'instrField')
    if (mark) {
      const instr = String(mark.attrs.instr)
      let next: string | undefined
      if (fieldKeyword(instr) === 'PAGE') {
        // resolve from the page the field's block sits on; when pagination
        // cannot place the block, keep the cached result instead of silently
        // writing a wrong value (BUG-1709)
        const blockPos = topLevelBlockPos(editor.state.doc, pos)
        const page = blockPos === null ? null : (pageOf()?.(blockPos) ?? null)
        if (page !== null) next = String(page)
      } else {
        const cached = resolvers.fieldValue(instr)
        if (cached) next = cached
      }
      if (next !== undefined && next !== node.text) {
        jobs.push({ from: pos, to: pos + node.nodeSize, text: next, marks: node.marks })
      }
      return
    }
    const ref = node.marks.find((m) => m.type.name === 'refField')
    if (ref) {
      // legacy references without a stored instruction fall back to the plain default
      const instr = String(ref.attrs.instr || ` REF ${ref.attrs.name} \\h `)
      const next = resolvers.refCache(instr, instr.includes('\\p') ? pageOf() : null)
      if (next !== null && next !== node.text) {
        jobs.push({ from: pos, to: pos + node.nodeSize, text: next, marks: node.marks })
      }
    }
  })
  return jobs
}

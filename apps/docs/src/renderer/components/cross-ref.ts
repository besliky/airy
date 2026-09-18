// Cross-reference (REF field) sources, types, and cache computation. Pure
// editor-doc helpers — no React — so tests cover them without the ribbon.
import type { Editor } from '@tiptap/core'
import { decodeEntities, type Block } from '@airy-office/docx-engine'

/** what a cross-reference displays (Word's "Insert reference to") */
export type CrossRefType = 'text' | 'page' | 'number'

/** which reference types each source kind offers (Word parity) */
export const CROSS_REF_TYPES: Record<CrossRefSourceKind, CrossRefType[]> = {
  heading: ['text', 'page', 'number'],
  bookmark: ['text', 'page'],
  caption: ['text', 'page', 'number'],
}

export type CrossRefSourceKind = 'heading' | 'bookmark' | 'caption'

/** one pickable cross-reference target */
export interface CrossRefSource {
  kind: CrossRefSourceKind
  /** main list label: heading text / bookmark name / "Figure 3 …" */
  label: string
  /** secondary column preview */
  preview: string
  /** heading outline level (1-6); others indent at level 1 */
  level: number
  /** existing anchor name (bookmark name, heading `_Toc…`, caption `_Ref…`); null = stamped on insert */
  anchor: string | null
  /** owning node position (anchor stamping / page lookup) */
  pos: number
  /** caption SEQ label ("Figure") and its document-order ordinal */
  seqLabel?: string
  seqNumber?: number
}

/** bookmark-paragraph text shown by a plain REF: Word renders the bookmarked
 * text; capped so one reference cannot paste a whole chapter inline */
export function refTextPreview(text: string): string {
  return text.trim().slice(0, 80)
}

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
function xmlOfNode(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  return blocks.find((b) => b.docxIndex === idx)?.originalXml ?? ''
}

/** the SEQ label of a protected caption paragraph, if it is one. Word often
 * splits one field instruction across several w:instrText runs (rsid seams),
 * so ALL of the paragraph's instruction fragments are joined before the SEQ
 * match — the same concatenation the docx reader does while parsing fields.
 * Reading only the first fragment turned "SEQ Fig|ure" into label "Fig" (a
 * wrong ordinal pool) or dropped the caption from the dialog entirely. */
function seqLabelOf(xml: string): string | null {
  const instr = (xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? [])
    .map((frag) => decodeEntities(frag.replace(/<[^>]*>/g, '')))
    .join('')
  const m = /^\s*SEQ\s+(\S+)/.exec(instr)
  return m ? m[1] : null
}

/** the first bookmark name inside a protected paragraph's XML (caption anchors) */
function xmlAnchorOf(xml: string): string | null {
  return /<w:bookmarkStart [^>]*w:name="([^"]+)"/.exec(xml)?.[1] ?? null
}

/** all bookmarks (visible + hidden node attrs), document order */
function collectBookmarks(editor: Editor): Array<{ name: string; pos: number; preview: string }> {
  const out: Array<{ name: string; pos: number; preview: string }> = []
  editor.state.doc.descendants((node, pos) => {
    const names = node.attrs?.bookmarks as string[] | null
    if (Array.isArray(names)) {
      for (const name of names) {
        out.push({ name, pos, preview: refTextPreview(node.textContent) })
      }
    }
    return !node.isLeaf
  })
  return out
}

/** Word's hidden heading anchor (`_Toc…`) already on the node, if any */
export function headingTocAnchor(node: { attrs?: Record<string, unknown> } | null): string | null {
  const hidden = node?.attrs?.hiddenBookmarks as string[] | null | undefined
  if (!Array.isArray(hidden)) return null
  return hidden.find((name) => /^_Toc\d+$/.test(name)) ?? null
}

/** every bookmark name in node attrs (visible + hidden) — link-dialog anchor uniqueness pool */
export function allBookmarkNames(doc: Editor['state']['doc']): Set<string> {
  const names = new Set<string>()
  doc.descendants((node) => {
    for (const attr of ['bookmarks', 'hiddenBookmarks'] as const) {
      const list = node.attrs?.[attr] as string[] | null | undefined
      if (Array.isArray(list)) for (const name of list) names.add(name)
    }
    return !node.isLeaf
  })
  return names
}

/** a fresh Word-style `_Toc` + 9-digit anchor name not colliding with `taken` */
export function uniqueTocAnchor(taken: Set<string>): string {
  return uniqueAnchor('_Toc', taken)
}

/**
 * The heading's cross-reference anchor: its existing hidden `_Toc…` bookmark,
 * or a fresh one stamped onto the node (hidden bookmarks re-emit as
 * w:bookmarkStart on save, so the REF field can resolve the heading).
 * Stamping is a document mutation, so it never happens on a read-only editor:
 * callers that reach here while editing is locked get null (an existing anchor
 * is still returned — reading it mutates nothing).
 */
export function ensureHeadingTocAnchor(editor: Editor, pos: number): string | null {
  const node = editor.state.doc.nodeAt(pos)
  if (!node) return null
  const existing = headingTocAnchor(node)
  if (existing) return existing
  if (!editor.isEditable) return null
  const name = uniqueTocAnchor(allBookmarkNames(editor.state.doc))
  const hidden = (node.attrs?.hiddenBookmarks as string[] | null) ?? []
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      hiddenBookmarks: [...hidden, name],
    }),
  )
  return name
}

/** every bookmark name in the document (node attrs + protected XML anchors) — anchor uniqueness pool */
export function allRefAnchorNames(doc: Editor['state']['doc'], blocks: Block[]): Set<string> {
  const names = new Set<string>()
  doc.descendants((node) => {
    for (const attr of ['bookmarks', 'hiddenBookmarks'] as const) {
      const list = node.attrs?.[attr] as string[] | null | undefined
      if (Array.isArray(list)) for (const name of list) names.add(name)
    }
    const anchor = xmlAnchorOf(xmlOfNode(node as never, blocks))
    if (anchor) names.add(anchor)
    return !node.isLeaf
  })
  return names
}

/** a fresh Word-style anchor (`_Ref…`/`_Toc…` + 9 digits) not colliding with `taken` */
export function uniqueAnchor(prefix: '_Ref' | '_Toc', taken: Set<string>): string {
  for (;;) {
    const name = `${prefix}${Math.floor(100000000 + Math.random() * 900000000)}`
    if (!taken.has(name)) return name
  }
}

/** cross-reference targets in document order: headings, captions (SEQ), bookmarks */
export function collectCrossRefSources(editor: Editor, blocks: Block[]): CrossRefSource[] {
  const out: CrossRefSource[] = []
  const seqOrdinals = new Map<string, number>()
  editor.state.doc.forEach((node, offset) => {
    const text = node.textContent.trim()
    if (node.type.name === 'docHeading' && text) {
      out.push({
        kind: 'heading',
        label: text,
        preview: refTextPreview(node.textContent),
        level: Number(node.attrs.level) || 1,
        anchor: headingTocAnchor(node),
        pos: offset,
      })
      return
    }
    if (node.type.name === 'docProtected') {
      const xml = xmlOfNode(node as never, blocks)
      const seqLabel = seqLabelOf(xml)
      if (seqLabel) {
        const n = (seqOrdinals.get(seqLabel) ?? 0) + 1
        seqOrdinals.set(seqLabel, n)
        const display = node.attrs.fieldDisplay as { left?: string } | null | undefined
        out.push({
          kind: 'caption',
          label: display?.left || `${seqLabel} ${n}`,
          preview: refTextPreview(node.textContent ?? ''),
          level: 1,
          anchor: xmlAnchorOf(xml),
          pos: offset,
          seqLabel,
          seqNumber: n,
        })
      }
    }
  })
  for (const b of collectBookmarks(editor)) {
    out.push({
      kind: 'bookmark',
      label: b.name,
      preview: b.preview,
      level: 1,
      anchor: b.name,
      pos: b.pos,
    })
  }
  return out
}

/**
 * Stamp a hidden `_Ref…` bookmark around the caption paragraph's content so a
 * REF field can resolve it (Word does this on caption insert). The node keeps
 * its XML via genXml (docxIndex cleared: the anchored XML now saves verbatim
 * from the node instead of the untouched original).
 */
export function ensureCaptionAnchor(
  editor: Editor,
  blocks: Block[],
  source: CrossRefSource,
): string | null {
  const node = editor.state.doc.nodeAt(source.pos)
  if (!node || node.type.name !== 'docProtected') return null
  const xml = xmlOfNode(node as never, blocks)
  if (xml === '') return null
  const existing = xmlAnchorOf(xml)
  if (existing) return existing
  if (!editor.isEditable) return null
  const name = uniqueAnchor('_Ref', allRefAnchorNames(editor.state.doc, blocks))
  const id = 1000000000 + Number(name.replace(/\D/g, '').slice(-9) || '0')
  const start = `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`
  const end = `<w:bookmarkEnd w:id="${id}"/>`
  // wrap the paragraph content: start after its pPr, end before </w:p>
  const withStart = xml.includes('</w:pPr>')
    ? xml.replace('</w:pPr>', `</w:pPr>${start}`)
    : start + xml
  const anchored = withStart.endsWith('</w:p>')
    ? `${withStart.slice(0, -'</w:p>'.length)}${end}</w:p>`
    : withStart + end
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(source.pos, undefined, {
      ...node.attrs,
      docxIndex: null,
      genXml: anchored,
    }),
  )
  return name
}

/** full REF instruction for a source and display type (written to w:instrText verbatim) */
export function crossRefInstr(anchor: string, type: CrossRefType): string {
  const sw = type === 'page' ? '\\p ' : type === 'number' ? '\\r ' : ''
  return ` REF ${anchor} ${sw}\\h `
}

/** cached display text for a freshly inserted reference (Word recomputes on field update) */
export function crossRefCache(
  source: CrossRefSource,
  type: CrossRefType,
  pageOf?: (pos: number) => number | null,
): string {
  if (type === 'page') {
    const page = pageOf?.(source.pos)
    // unknown until pagination: blank cache, filled by F9 / Word
    return page !== undefined && page !== null ? String(page) : ' '
  }
  if (type === 'number') {
    if (source.kind === 'caption') return String(source.seqNumber ?? 1)
    // heading: REF \r shows the outline number — only computable when the
    // heading text itself carries it (numbered style prefix)
    const m = /^(\d+(?:\.\d+)*)[.、\s]/.exec(source.label)
    return m ? m[1] : ' '
  }
  return source.kind === 'bookmark' ? source.preview || source.label : source.label
}

/** position of the node carrying `name` (node bookmark attrs or a protected-XML anchor) */
export function findAnchorPos(
  doc: Editor['state']['doc'],
  blocks: Block[],
  name: string,
): number | null {
  let found: number | null = null
  doc.descendants((node, pos) => {
    if (found !== null) return
    for (const attr of ['bookmarks', 'hiddenBookmarks'] as const) {
      const list = node.attrs?.[attr] as string[] | null | undefined
      if (Array.isArray(list) && list.includes(name)) {
        found = pos
        return
      }
    }
    if (
      node.type.name === 'docProtected' &&
      xmlOfNode(node as never, blocks).includes(`w:name="${name}"`)
    ) {
      found = pos
    }
  })
  return found
}

/** current SEQ ordinal of the caption at `pos` (document-order count per label) */
function seqNumberAt(
  doc: Editor['state']['doc'],
  blocks: Block[],
  pos: number,
  seqLabel: string,
): string | null {
  const target = doc.nodeAt(pos)
  if (!target) return null
  let n = 0
  doc.forEach((node, offset) => {
    if (offset > pos) return
    if (node.type.name !== 'docProtected') return
    if (seqLabelOf(xmlOfNode(node as never, blocks)) === seqLabel) n += 1
  })
  return n > 0 ? String(n) : null
}

/**
 * F9 cache recompute for one REF instruction. Returns the new display text, or
 * null when the target is gone / the value is not computable locally (page
 * number before pagination): the existing cache stays untouched then.
 */
export function refCacheOf(
  editor: Editor,
  blocks: Block[],
  instr: string,
  pageOf?: ((pos: number) => number | null) | null,
): string | null {
  const m = /^\s*REF\s+(?:"([^"]+)"|(\S+))/.exec(instr)
  if (!m) return null
  const name = m[1] ?? m[2]
  const pos = findAnchorPos(editor.state.doc, blocks, name)
  if (pos === null) return null
  if (instr.includes('\\p')) {
    const page = pageOf?.(pos)
    return page !== undefined && page !== null ? String(page) : null
  }
  if (instr.includes('\\r')) {
    const node = editor.state.doc.nodeAt(pos)
    if (node?.type.name === 'docProtected') {
      const label = seqLabelOf(xmlOfNode(node as never, blocks))
      return label ? seqNumberAt(editor.state.doc, blocks, pos, label) : null
    }
    const hm = /^(\d+(?:\.\d+)*)[.、\s]/.exec(node?.textContent.trim() ?? '')
    return hm ? hm[1] : null
  }
  const node = editor.state.doc.nodeAt(pos)
  const text = refTextPreview(node?.textContent ?? '')
  return text || name
}

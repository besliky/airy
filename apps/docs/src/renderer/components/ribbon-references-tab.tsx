import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import {
  bibliographyLine,
  citationText,
  decodeEntities,
  generateCaptionXml,
  generateIndexFieldXml,
  generateTocFieldXml,
  parseTocInstruction,
  tocEntryHidesPage,
  type Block,
  type SourceInfo,
  type StyleInfo,
  type TocEntry,
  type TocFieldOptions,
} from '@airy-office/docx-engine'
import { Dropdown, useModalDialog } from '@airy-office/ui'
import { PromptModal } from './PromptModal'
import { allRefAnchorNames, collectCrossRefSources, uniqueAnchor } from './cross-ref'
import { collectHeadings } from '../editor/headings'
import { t, useI18n, type StringKey } from '../i18n/locale'
import {
  IconBook,
  IconCaption,
  IconCaret,
  IconCitation,
  IconEndnote,
  IconFootnote,
  IconGear,
  IconIndex,
  IconRefresh,
  IconToc,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, toggleDropdown } from './ribbon-tabs'

function collectTocEntries(editor: Editor): TocEntry[] {
  return collectHeadings(editor.state.doc).map(({ level, text }) => ({ level, text }))
}

/** Heading entries + real page numbers (headingPages and collectTocEntries share document order) */
function collectTocEntriesWithPages(
  editor: Editor,
  headingPages?: () => number[] | null,
): TocEntry[] {
  const entries = collectTocEntries(editor)
  const pages = headingPages?.()
  if (pages && pages.length === entries.length) {
    return entries.map((e, i) => ({ ...e, pageNo: pages[i] }))
  }
  return entries
}

/** one `StyleName,level` pair of a \t source-styles spec ("A,1,B,2") */
export function parseTocStyleSpec(spec: string): Array<{ name: string; level: number }> {
  const parts = spec.split(',').map((p) => p.trim())
  const out: Array<{ name: string; level: number }> = []
  for (let i = 0; i + 1 < parts.length; i += 2) {
    const name = parts[i]
    const level = parseInt(parts[i + 1], 10)
    if (name && Number.isFinite(level) && level >= 1 && level <= 9)
      out.push({ name, level: Math.round(level) })
  }
  return out
}

/**
 * Entries from the \t source-styles mapping: paragraphs whose style matches a
 * `StyleName,level` pair of the spec, at that level, in document order with
 * real page numbers when pagination is available. The instruction carries
 * style NAMES while a paragraph's pStyle carries styleIds — the document's
 * parsed styles bridge the two (BUG-1012); a spec token that is itself a
 * styleId (some producers write ids into \t) still matches directly.
 */
export function collectStyledTocEntries(
  editor: Editor,
  styles: Map<string, StyleInfo> | undefined,
  spec: string,
  anchorPage?: (pos: number) => number | null,
): Array<TocEntry & { pos: number }> {
  const pairs = parseTocStyleSpec(spec)
  if (pairs.length === 0) return []
  const norm = (s: string) => s.trim().toLowerCase()
  const levelOf = new Map<string, number>()
  if (styles)
    for (const info of styles.values()) {
      if (info.type !== 'paragraph') continue
      for (const { name, level } of pairs) {
        if (norm(info.name) === norm(name) || norm(info.styleId) === norm(name))
          if (!levelOf.has(info.styleId)) levelOf.set(info.styleId, level)
      }
    }
  for (const { name, level } of pairs) if (!levelOf.has(name)) levelOf.set(name, level)
  const out: Array<TocEntry & { pos: number }> = []
  editor.state.doc.forEach((node, offset) => {
    const styleId = node.attrs?.styleId
    const level = typeof styleId === 'string' ? levelOf.get(styleId) : undefined
    if (level === undefined) return
    const text = node.textContent.trim()
    if (!text) return
    const page = anchorPage?.(offset)
    out.push({
      level,
      text,
      ...(page !== null && page !== undefined ? { pageNo: page } : {}),
      pos: offset,
    })
  })
  return out
}

/**
 * TOC entries for the parsed/dialog options, in document order: the \o
 * heading range unions with the \t style-mapped paragraphs (Word's semantics
 * when an instruction carries both); a styles-only TOC (\t without \o)
 * collects the mapped styles alone (BUG-1012). A paragraph collected by both
 * sources keeps its heading entry.
 */
export function collectTocEntriesForOptions(
  editor: Editor,
  options: TocFieldOptions,
  ctx: {
    headingPages?: () => number[] | null
    anchorPage?: (pos: number) => number | null
    styles?: Map<string, StyleInfo>
  } = {},
): TocEntry[] {
  const byPos = new Map<number, TocEntry>()
  if (!options.styles || options.levels !== undefined) {
    const pages = ctx.headingPages?.()
    collectHeadings(editor.state.doc).forEach((h, i) => {
      if (options.levels !== undefined && h.level > options.levels) return
      byPos.set(h.pos, {
        level: h.level,
        text: h.text,
        ...(pages && pages[i] !== undefined ? { pageNo: pages[i] } : {}),
      })
    })
  }
  if (options.styles) {
    for (const entry of collectStyledTocEntries(
      editor,
      ctx.styles,
      options.styles,
      ctx.anchorPage,
    )) {
      const { pos, ...rest } = entry
      if (!byPos.has(pos)) byPos.set(pos, rest)
    }
  }
  return [...byPos.entries()].sort((a, b) => a[0] - b[0]).map(([, e]) => e)
}

/**
 * Table-of-figures entries from SEQ captions of one label ("Figure"), in
 * document order. Each entry keeps the caption's `_Ref…` anchor (click-to-jump)
 * and its real page number when pagination is available. `labels` carries the
 * canonical identifier plus (for legacy documents) the translated SEQ word the
 * caption may have been authored with — see CAPTION_LABELS (UX-1011).
 */
export function collectTofEntries(
  editor: Editor,
  blocks: Block[],
  labels: string[] | string,
  anchorPage?: (pos: number) => number | null,
): TocEntry[] {
  const wanted = Array.isArray(labels) ? labels : [labels]
  return collectCrossRefSources(editor, blocks)
    .filter((s) => s.kind === 'caption' && s.seqLabel !== undefined && wanted.includes(s.seqLabel))
    .map((s) => ({
      level: 1,
      text: s.label,
      pageNo: anchorPage?.(s.pos) ?? undefined,
      ...(s.anchor ? { anchor: s.anchor } : {}),
    }))
}

/** PM nodes for a freshly generated TOC/TOF field (one docProtected per line) */
function tocFieldNodes(
  entries: TocEntry[],
  options: TocFieldOptions = {},
  nodeLabel = t('ribbonTocFieldLabel'),
): Array<Record<string, unknown>> {
  return generateTocFieldXml(entries, options).map((xml, i) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: nodeLabel,
      genXml: xml,
      fieldDisplay: {
        kind: 'tocLine',
        left: entries[i].text,
        // a ranged \n hides the page number per entry level (BUG-1012)
        right:
          tocEntryHidesPage(options, entries[i].level) || entries[i].pageNo === undefined
            ? ''
            : String(entries[i].pageNo),
        level: entries[i].level,
        ...(tocEntryHidesPage(options, entries[i].level) ? { noPage: true } : {}),
        ...(entries[i].anchor ? { anchor: entries[i].anchor } : {}),
      },
    },
  }))
}

const PAGE_BREAK_PARAGRAPH_XML = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
function xmlOfNode(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  return blocks.find((b) => b.docxIndex === idx)?.originalXml ?? ''
}

/** all field instruction fragments of a paragraph joined (Word splits across runs) */
function instrTextOf(xml: string): string {
  return (xml.match(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g) ?? [])
    .map((frag) => decodeEntities(frag.replace(/<[^>]*>/g, '')))
    .join('')
}

export type TocUpdateResult = 'updated' | 'missing' | 'no-entries'

/**
 * The identifiers a ToF update may collect for an authored `\c` identifier:
 * the identifier itself plus — when it is (or aliases) a known caption label —
 * the canonical id and the current locale's word for it. Legacy documents
 * authored before UX-1011 stored the translated word in their SEQ
 * instructions; without the aliases the F9/Update path would rebuild a table
 * of figures from nothing while Insert Table of Figures finds the captions
 * (BUG-1110), and mixed old/new captions would split into two independent
 * SEQ series.
 */
function tofLabelAliases(id: string): string[] {
  const out = new Set([id])
  for (const { id: canonical, key } of CAPTION_LABELS) {
    const translated = t(key)
    if (id === canonical || id === translated) {
      out.add(canonical)
      if (translated && translated !== canonical) out.add(translated)
    }
  }
  return [...out]
}

/**
 * The section properties riding inside the TOC/ToF region's paragraphs. A
 * section's properties live in the pPr of its last paragraph: the trailing
 * one on the region's last field entry (BUG-1003), plus any CONTINUOUS breaks
 * inside the region — a multi-column TOC is a chain of sections, one per
 * column. Deleting the region without re-emitting them destroyed every
 * mid-region section (BUG-1107); ALL of them are collected here, in document
 * order, and re-attached by attachRegionSectPrs.
 */
const SECT_PR_ALL_RE = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g

/** inject a sectPr fragment into a generated paragraph's pPr (CT_PPr order: after rPr, at the end) */
function attachSectPr(xml: string, sectPr: string): string {
  return xml.includes('</w:pPr>')
    ? xml.replace('</w:pPr>', `${sectPr}</w:pPr>`)
    : xml.replace(/^<w:p(\s[^>]*)?>/, (open) => `${open}<w:pPr>${sectPr}</w:pPr>`)
}

/** an empty paragraph carrying a sectPr — a Word section break paragraph
 *  (the parser labels content-less sectPr paragraphs the same way) */
function sectionBreakNode(sectPr: string): Record<string, unknown> {
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: 'Section break paragraph',
      genXml: `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`,
    },
  }
}

/**
 * Re-emit the region's collected sectPr fragments onto the regenerated
 * paragraphs (BUG-1107). Each break ends a section, so they keep their
 * document order and the trailing one stays on the last paragraph
 * (BUG-1003); earlier ones spread over intermediate paragraphs — a
 * multi-column TOC keeps its column boundaries. More breaks than
 * regenerated paragraphs (the degenerate empty-sections case) keep their
 * own break paragraphs, in order, after the paragraph holding the
 * previous break.
 */
function attachRegionSectPrs(nodes: Array<Record<string, unknown>>, sectPrs: string[]): void {
  if (sectPrs.length === 0 || nodes.length === 0) return
  const n = nodes.length
  let lastIdx = -1
  const surplus: Array<{ afterIdx: number; sectPr: string }> = []
  for (let i = 0; i < sectPrs.length; i++) {
    const target = Math.max(lastIdx + 1, Math.floor(((i + 1) * n) / sectPrs.length) - 1)
    if (target <= n - 1) {
      lastIdx = target
      const node = nodes[target] as { attrs: Record<string, unknown> }
      node.attrs.genXml = attachSectPr(String(node.attrs.genXml), sectPrs[i])
    } else surplus.push({ afterIdx: lastIdx, sectPr: sectPrs[i] })
  }
  for (let s = surplus.length - 1; s >= 0; s--) {
    nodes.splice(surplus[s].afterIdx + 1, 0, sectionBreakNode(surplus[s].sectPr))
  }
}

/**
 * Find every TOC/TOF field region (field begin ... matching end, tracked by
 * fldChar depth across top-level blocks) and replace ALL of them with
 * regenerated dirty fields in one transaction (BUG-1011: documents can carry
 * a TOC and a table of figures — only the first used to update). Each
 * authored instruction is parsed first, so an update keeps the chosen
 * switches (level range, \n and its per-level range, \t styles) and rebuilds
 * a table of figures (\c label) from SEQ captions instead of headings.
 * `styles` is the document's parsed style map: \t matches style NAMES while
 * paragraphs carry styleIds, so without it a style-mapped TOC would fall
 * back to heading entries (BUG-1012).
 */
export function updateTocField(
  editor: Editor,
  blocks: Block[],
  headingPages?: () => number[] | null,
  anchorPage?: (pos: number) => number | null,
  styles?: Map<string, StyleInfo>,
  opts: { silent?: boolean } = {},
): TocUpdateResult {
  // every TOC/TOF field region (begin ... matching end, tracked by fldChar
  // depth across top-level blocks) — an update rebuilds ALL of them, like
  // Word's update-fields pass, instead of only the first (BUG-1011)
  const doc = editor.state.doc
  const regions: Array<{
    from: number
    to: number
    instr: string
    keepPageBreak: boolean
    sectPrs: string[]
  }> = []
  let from = -1
  let instr = ''
  let depth = 0
  let regionSectPrs: string[] = []
  doc.forEach((node, offset) => {
    const xml = xmlOfNode(node as never, blocks)
    if (from === -1) {
      const joined = instrTextOf(xml)
      if (!/^\s*TOC[\s\\]/.test(joined)) return
      from = offset
      instr = joined
      depth = 0
      regionSectPrs = []
    }
    // harvest EVERY sectPr of the region's nodes in document order: the
    // trailing one plus mid-region continuous breaks (BUG-1107)
    regionSectPrs.push(...(xml.match(SECT_PR_ALL_RE) ?? []))
    depth += (xml.match(/w:fldCharType="begin"/g) ?? []).length
    depth -= (xml.match(/w:fldCharType="end"/g) ?? []).length
    if (depth <= 0) {
      regions.push({
        from,
        to: offset + node.nodeSize,
        instr,
        keepPageBreak: /<w:br\s[^>]*w:type="page"/.test(xml),
        sectPrs: regionSectPrs,
      })
      from = -1
    }
  })

  if (regions.length === 0) {
    if (!opts.silent) window.alert(t('ribbonTocNotFound'))
    return 'missing'
  }

  // plan every region against the same doc snapshot, then apply the
  // replacements back-to-front in ONE transaction (positions of earlier
  // regions stay valid; one undo step for the whole update)
  let updated = 0
  const planned: Array<{
    region: (typeof regions)[number]
    nodes: ReturnType<typeof tocFieldNodes>
  }> = []
  for (const region of regions) {
    const options = parseTocInstruction(region.instr)
    // BUG-1110: the update path collects ToF entries under the same
    // canonical+locale alias set the Insert dialog uses (seqAliases/UX-1011);
    // BUG-1012: a \t style-mapped TOC collects entries by style (union with
    // the \o heading range when the instruction carries both)
    const entries = options.seqIdentifier
      ? collectTofEntries(editor, blocks, tofLabelAliases(options.seqIdentifier), anchorPage)
      : collectTocEntriesForOptions(editor, options, { headingPages, anchorPage, styles })
    if (entries.length === 0) {
      continue // empty field: skipped, reported in the aggregate below
    }
    const nodes = tocFieldNodes(
      entries,
      options,
      t(options.seqIdentifier ? 'refsTofFieldLabel' : 'ribbonTocFieldLabel'),
    )
    if (region.keepPageBreak) {
      nodes.push({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: t('ribbonPageBreak'),
          genXml: PAGE_BREAK_PARAGRAPH_XML,
          fieldDisplay: { kind: 'pageBreak' },
        },
      })
    }
    // the region's section breaks move back onto the regenerated paragraphs:
    // the trailing one onto the last entry (BUG-1003), continuous mid-region
    // breaks onto intermediate entries in document order (BUG-1107)
    attachRegionSectPrs(nodes, region.sectPrs)
    planned.push({ region, nodes })
    updated += 1
  }

  if (updated === 0) {
    const first = parseTocInstruction(regions[0].instr)
    if (!opts.silent)
      window.alert(t(first.seqIdentifier ? 'refsTofNoCaptions' : 'ribbonTocNoHeadings'))
    return 'no-entries'
  }

  const chain = editor.chain().focus()
  for (let i = planned.length - 1; i >= 0; i--) {
    const { region, nodes } = planned[i]
    chain
      .deleteRange({ from: region.from, to: region.to })
      .insertContentAt(region.from, nodes as never)
  }
  chain.run()
  return 'updated'
}

/**
 * Caption labels: the `id` is the language-independent SEQ / TOC \c identifier
 * stored in the document; the `key` translates the visible word. Captions and
 * tables of figures keep working across UI-language switches because matching
 * and numbering run on the id, not on the translated string (UX-1011).
 */
const CAPTION_LABELS = [
  { id: 'Figure', key: 'ribbonCaptionFigure' },
  { id: 'Table', key: 'ribbonCaptionTable' },
  { id: 'Equation', key: 'ribbonCaptionEquation' },
] as const

/**
 * The identifiers a caption of label `id` may carry in this document: the
 * canonical id, plus the current locale's word for it — captions authored
 * before UX-1011 (or by a differently-localized writer) stored the translated
 * word in their SEQ instruction.
 */
function seqAliases(id: string, t: (key: StringKey) => string): string[] {
  const entry = CAPTION_LABELS.find((l) => l.id === id)
  const translated = entry ? t(entry.key) : null
  return translated && translated !== id ? [id, translated] : [id]
}

/** the visible (translated) label word of a canonical caption id */
function captionDisplayLabel(id: string, t: (key: StringKey) => string): string {
  const entry = CAPTION_LABELS.find((l) => l.id === id)
  return entry ? t(entry.key) : id
}

/** TOC options dialog (Word's Table of Contents options): level range,
 *  page numbers + hyperlinks switches, optional source styles (\t).
 *  Exported for its DOM tests (empty states stay inline, UX-1010). */
export function TocOptionsModal({
  editor,
  headingPages,
  anchorPage,
  docStyles,
  onClose,
}: {
  editor: Editor
  headingPages?: () => number[] | null
  anchorPage?: (pos: number) => number | null
  /** document style map (name→styleId bridging for the \t spec, BUG-1012) */
  docStyles?: Map<string, StyleInfo>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [levels, setLevels] = useState('3')
  const [showPages, setShowPages] = useState(true)
  const [hyperlinks, setHyperlinks] = useState(true)
  const [styles, setStyles] = useState('')
  // inline empty-state message (UX-1010): no blocking window.alert over the dialog
  const [error, setError] = useState<StringKey | null>(null)
  const dialog = useModalDialog(onClose)

  const levelCount = Math.min(Math.max(parseInt(levels, 10) || 3, 1), 9)
  const insert = () => {
    const options: TocFieldOptions = {
      levels: levelCount,
      ...(showPages ? {} : { hidePageNumbers: true }),
      ...(hyperlinks ? {} : { hyperlinks: false }),
      ...(styles.trim() ? { styles: styles.trim() } : {}),
    }
    // \t style-mapped paragraphs union with the heading range (BUG-1012)
    const entries = collectTocEntriesForOptions(editor, options, {
      headingPages,
      anchorPage,
      styles: docStyles,
    })
    if (entries.length === 0) {
      setError('ribbonTocNoHeadings')
      return
    }
    editor
      .chain()
      .focus()
      .insertContent(tocFieldNodes(entries, options) as never)
      .run()
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('refsTocOptionsTitle')}</h2>
        <label>
          {t('refsTocLevels')}
          <Dropdown
            value={levels}
            ariaLabel={t('refsTocLevels')}
            options={Array.from({ length: 9 }, (_, i) => String(i + 1)).map((v) => ({
              value: v,
              label: v,
            }))}
            onPick={setLevels}
          />
        </label>
        <label className="font-check">
          <input
            type="checkbox"
            checked={showPages}
            onChange={(e) => setShowPages(e.target.checked)}
          />
          {t('refsTocPageNumbers')}
        </label>
        <label className="font-check">
          <input
            type="checkbox"
            checked={hyperlinks}
            onChange={(e) => setHyperlinks(e.target.checked)}
          />
          {t('refsTocHyperlinks')}
        </label>
        <label>
          {t('refsTocStyles')}
          <input
            value={styles}
            onChange={(e) => setStyles(e.target.value)}
            placeholder={t('refsTocStylesPh')}
          />
        </label>
        {error && (
          <p className="modal-error" role="alert">
            {t(error)}
          </p>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Table of Figures dialog: pick the caption label the TOC \c field collects.
 *  Exported for its DOM tests (empty states stay inline, UX-1010). */
export function TofModal({
  editor,
  blocks,
  anchorPage,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  anchorPage?: (pos: number) => number | null
  onClose: () => void
}) {
  const { t } = useI18n()
  // the picked canonical label id — the dropdown shows translated words
  const [labelId, setLabelId] = useState<string>(CAPTION_LABELS[0].id)
  // inline empty-state message (UX-1010): no blocking window.alert over the dialog
  const [error, setError] = useState<StringKey | null>(null)
  const dialog = useModalDialog(onClose)

  const insert = () => {
    const entries = collectTofEntries(editor, blocks, seqAliases(labelId, t), anchorPage)
    if (entries.length === 0) {
      setError('refsTofNoCaptions')
      return
    }
    editor
      .chain()
      .focus()
      .insertContent(
        tocFieldNodes(entries, { seqIdentifier: labelId }, t('refsTofFieldLabel')) as never,
      )
      .run()
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('refsTofTitle')}</h2>
        <label>
          {t('refsTofLabel')}
          <Dropdown
            value={labelId}
            ariaLabel={t('refsTofLabel')}
            options={CAPTION_LABELS.map(({ id, key }) => ({ value: id, label: t(key) }))}
            onPick={setLabelId}
          />
        </label>
        {error && (
          <p className="modal-error" role="alert">
            {t(error)}
          </p>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Caption dialog: label + optional description, numbered by existing SEQ fields */
function CaptionModal({
  editor,
  blocks,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  onClose: () => void
}) {
  const { t } = useI18n()
  // the picked canonical label id; the visible word comes from the locale
  const [labelId, setLabelId] = useState<string>(CAPTION_LABELS[0].id)
  const [text, setText] = useState('')

  const nextNumber = (id: string): number => {
    // count SEQ fields of every identifier this label may carry (canonical +
    // legacy translated word, UX-1011)
    const alternation = seqAliases(id, t)
      .map((a) => a.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
      .join('|')
    const re = new RegExp(`SEQ\\s+(?:${alternation})[\\s\\\\]`)
    let count = 0
    editor.state.doc.forEach((node) => {
      if (node.type.name !== 'docProtected') return
      if (re.test(xmlOfNode(node as never, blocks))) count += 1
    })
    return count + 1
  }

  const insert = () => {
    const displayLabel = captionDisplayLabel(labelId, t)
    const number = nextNumber(labelId)
    // hidden _Ref anchor wraps the SEQ field so cross-references can target this caption
    const anchor = uniqueAnchor('_Ref', allRefAnchorNames(editor.state.doc, blocks))
    // the SEQ instruction carries the canonical id; only the visible prefix
    // is the translated word (language-independent matching, UX-1011)
    const xml = generateCaptionXml(labelId, number, text.trim(), anchor, displayLabel)
    const display = `${displayLabel} ${number}${text.trim() ? ` ${text.trim()}` : ''}`
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: t('ribbonCaption'),
          genXml: xml,
          fieldDisplay: { kind: 'text', left: display },
        },
      } as never)
      .run()
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonCaptionInsertTitle')}</h2>
        <label>
          {t('ribbonCaptionLabel')}
          <Dropdown
            value={labelId}
            ariaLabel={t('ribbonCaptionLabel')}
            options={CAPTION_LABELS.map(({ id, key }) => ({ value: id, label: t(key) }))}
            onPick={setLabelId}
          />
        </label>
        <label>
          {t('ribbonCaption')}
          <input
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder={t('ribbonCaptionPh', {
              label: captionDisplayLabel(labelId, t),
              n: nextNumber(labelId),
            })}
            onKeyDown={(e) => e.key === 'Enter' && insert()}
          />
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

const SOURCE_TYPES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'Book', nameKey: 'ribbonSourceBook' },
  { key: 'JournalArticle', nameKey: 'ribbonSourceJournal' },
  { key: 'InternetSite', nameKey: 'ribbonSourceWebsite' },
  { key: 'Report', nameKey: 'ribbonSourceReport' },
  { key: 'Misc', nameKey: 'ribbonSourceMisc' },
]

/** Create Source dialog for the citation manager */
function SourceModal({
  sources,
  onAdd,
  onClose,
}: {
  sources: SourceInfo[]
  onAdd: (source: SourceInfo) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [type, setType] = useState('Book')
  const [author, setAuthor] = useState('')
  const [title, setTitle] = useState('')
  const [year, setYear] = useState('')
  const [publisher, setPublisher] = useState('')
  const [url, setUrl] = useState('')

  const add = () => {
    if (!title.trim()) return
    // unique tag derived from author/title initials + year
    const base = (author.trim() || title.trim()).replace(/[\s,，、]+/g, '').slice(0, 8) || 'Src'
    let tag = `${base}${year.trim()}`
    let n = 1
    while (sources.some((s) => s.tag === tag)) tag = `${base}${year.trim()}_${++n}`
    onAdd({
      tag,
      type,
      author: author.trim(),
      title: title.trim(),
      year: year.trim(),
      ...(publisher.trim() ? { publisher: publisher.trim() } : {}),
      ...(url.trim() ? { url: url.trim() } : {}),
    })
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonSourceCreateTitle')}</h2>
        <label>
          {t('ribbonSourceType')}
          <Dropdown
            value={type}
            ariaLabel={t('ribbonSourceType')}
            options={SOURCE_TYPES.map((s) => ({ value: s.key, label: t(s.nameKey) }))}
            onPick={setType}
          />
        </label>
        <label>
          {t('ribbonSourceAuthor')}
          <input
            value={author}
            onChange={(e) => setAuthor(e.target.value)}
            placeholder={t('ribbonSourceAuthorPh')}
          />
        </label>
        <label>
          {t('ribbonSourceTitle')}
          <input
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder={t('ribbonSourceTitlePh')}
          />
        </label>
        <label>
          {t('ribbonSourceYear')}
          <input value={year} onChange={(e) => setYear(e.target.value)} placeholder="2026" />
        </label>
        <label>
          {type === 'JournalArticle'
            ? t('ribbonSourceJournalName')
            : type === 'InternetSite'
              ? t('ribbonSourceSiteName')
              : t('ribbonSourcePublisher')}
          <input value={publisher} onChange={(e) => setPublisher(e.target.value)} />
        </label>
        {type === 'InternetSite' && (
          <label>
            URL
            <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://…" />
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!title.trim()} onClick={add}>
            {t('ribbonAdd')}
          </button>
        </div>
      </div>
    </div>
  )
}

interface ReferencesTabProps extends TabProps {
  blocks: Block[]
  /** document style map: \t source-styles matching by name (BUG-1012) */
  docStyles?: Map<string, StyleInfo>
  onInsertNote: (kind: 'footnote' | 'endnote') => void
  /** open the Footnote and Endnote options dialog (numbering format, conversion) */
  onNoteOptions?: () => void
  /** Next / Previous note reference navigation (Word's Next Footnote) */
  onNavigateNote?: (dir: 1 | -1) => void
  sources: SourceInfo[]
  onAddSource: (source: SourceInfo) => void
  /** TOC page-number backfill: docHeadings in document order → real page numbers */
  headingPages?: () => number[] | null
  /** ToF page numbers: displayed page of any node position from live pagination */
  anchorPage?: (pos: number) => number | null
}

export function ReferencesTab({
  editor,
  hasDoc,
  blocks,
  docStyles,
  dropdown,
  setDropdown,
  onInsertNote,
  onNoteOptions,
  onNavigateNote,
  sources,
  onAddSource,
  headingPages,
  anchorPage,
}: ReferencesTabProps) {
  const { t } = useI18n()
  const [captionOpen, setCaptionOpen] = useState(false)
  const [sourceOpen, setSourceOpen] = useState(false)
  const [tocOptionsOpen, setTocOptionsOpen] = useState(false)
  const [tofOpen, setTofOpen] = useState(false)

  const insertToc = () => {
    const entries = collectTocEntriesWithPages(editor, headingPages)
    if (entries.length === 0) {
      window.alert(t('ribbonTocNoHeadings'))
      return
    }
    editor
      .chain()
      .focus()
      .insertContent(tocFieldNodes(entries) as never)
      .run()
  }

  const insertCitation = (source: SourceInfo) => {
    editor
      .chain()
      .focus()
      .insertContent({ type: 'text', text: citationText(source) } as never)
      .run()
    setDropdown(() => null)
  }

  const insertBibliography = () => {
    if (sources.length === 0) {
      window.alert(t('ribbonNoSources'))
      return
    }
    const nodes: Array<Record<string, unknown>> = [
      {
        type: 'docHeading',
        attrs: { docxIndex: null, styleId: null, aiChanged: false, level: 1 },
        content: [{ type: 'text', text: t('ribbonBibliographyHeading') }],
      },
      ...sources.map((s) => ({
        type: 'docParagraph',
        attrs: { docxIndex: null, styleId: null, aiChanged: false },
        content: [{ type: 'text', text: bibliographyLine(s) }],
      })),
    ]
    editor
      .chain()
      .focus()
      .insertContent(nodes as never)
      .run()
    setDropdown(() => null)
  }

  const [xePrompt, setXePrompt] = useState<{ initial: string; at: number } | null>(null)

  const markIndexEntry = () => {
    const { from, to } = editor.state.selection
    setXePrompt({ initial: editor.state.doc.textBetween(from, to, ' ').trim(), at: to })
  }

  const submitIndexEntry = (term: string) => {
    if (!xePrompt) return
    editor
      .chain()
      .focus()
      .insertContentAt(Math.min(xePrompt.at, editor.state.doc.content.size), {
        type: 'docXeMark',
        attrs: { term },
      } as never)
      .run()
  }

  const insertIndex = () => {
    const terms: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docXeMark') terms.push(String(node.attrs.term))
    })
    if (terms.length === 0) {
      window.alert(t('ribbonNoIndexEntries'))
      return
    }
    const unique = [...new Set(terms.map((t) => t.trim()).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, 'zh-CN'),
    )
    const nodes = generateIndexFieldXml(unique).map((xml, i) => ({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonIndexFieldLabel'),
        genXml: xml,
        fieldDisplay: { kind: 'tocLine', left: unique[i], right: '', level: 1 },
      },
    }))
    editor
      .chain()
      .focus()
      .insertContent(nodes as never)
      .run()
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonTocTip')}
              onClick={() => toggleDropdown(setDropdown, 'tocMenu')}
            >
              <span className="rb-big-icon">
                <IconToc size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonToc')}</span>
            </button>
            {dropdown === 'tocMenu' && (
              <div data-rb-panel="" className="layout-menu">
                <button
                  onClick={() => {
                    insertToc()
                    setDropdown(() => null)
                  }}
                >
                  {t('refsTocMenuAuto')}
                </button>
                <button
                  onClick={() => {
                    setTocOptionsOpen(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('refsTocMenuCustom')}
                </button>
                <button
                  onClick={() => {
                    setTofOpen(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('refsTocMenuTof')}
                </button>
              </div>
            )}
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonTocUpdateTip')}
            onClick={() => updateTocField(editor, blocks, headingPages, anchorPage, docStyles)}
          >
            <span className="rb-big-icon">
              <IconRefresh size={BIG} />
            </span>
            <span>{t('ribbonTocUpdate')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonToc')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonFootnoteTip')}
            onClick={() => onInsertNote('footnote')}
          >
            <span className="rb-big-icon">
              <IconFootnote size={BIG} />
            </span>
            <span>{t('ribbonFootnote')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonEndnoteTip')}
            onClick={() => onInsertNote('endnote')}
          >
            <span className="rb-big-icon">
              <IconEndnote size={BIG} />
            </span>
            <span>{t('ribbonEndnote')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('refsNoteMenuTip')}
              onClick={() => toggleDropdown(setDropdown, 'noteoptions')}
            >
              <span className="rb-big-icon">
                <IconGear size={BIG} />
                <IconCaret />
              </span>
              <span>{t('refsNoteMenu')}</span>
            </button>
            {dropdown === 'noteoptions' && (
              <div data-rb-panel="" className="layout-menu">
                <button
                  onClick={() => {
                    onNoteOptions?.()
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('refsNoteOptionsTitle')}</b>
                </button>
                <button
                  onClick={() => {
                    onNavigateNote?.(1)
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('refsNextNote')}</b>
                </button>
                <button
                  onClick={() => {
                    onNavigateNote?.(-1)
                    setDropdown(() => null)
                  }}
                >
                  <b>{t('refsPrevNote')}</b>
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupFootnotes')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonCitationTip')}
              onClick={() => toggleDropdown(setDropdown, 'citation')}
            >
              <span className="rb-big-icon">
                <IconCitation size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonCitation')}</span>
            </button>
            {dropdown === 'citation' && (
              <div data-rb-panel="" className="layout-menu">
                {sources.map((s) => (
                  <button key={s.tag} data-tip={s.title} onClick={() => insertCitation(s)}>
                    {citationText(s)} {s.title.slice(0, 12)}
                  </button>
                ))}
                <button
                  onClick={() => {
                    setSourceOpen(true)
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonAddNewSource')}
                </button>
              </div>
            )}
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonBibliographyTip')}
            onClick={insertBibliography}
          >
            <span className="rb-big-icon">
              <IconBook size={BIG} />
            </span>
            <span>{t('ribbonBibliography')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonCaptionTip')}
            onClick={() => setCaptionOpen(true)}
          >
            <span className="rb-big-icon">
              <IconCaption size={BIG} />
            </span>
            <span>{t('ribbonCaption')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonIndexTip')}
              onClick={() => toggleDropdown(setDropdown, 'index')}
            >
              <span className="rb-big-icon">
                <IconIndex size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonIndex')}</span>
            </button>
            {dropdown === 'index' && (
              <div data-rb-panel="" className="layout-menu">
                <button
                  onClick={() => {
                    markIndexEntry()
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonMarkEntry')}
                </button>
                <button
                  onClick={() => {
                    insertIndex()
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonInsertIndex')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupCitationsIndex')}</div>
      </div>

      {captionOpen && (
        <CaptionModal editor={editor} blocks={blocks} onClose={() => setCaptionOpen(false)} />
      )}
      {sourceOpen && (
        <SourceModal sources={sources} onAdd={onAddSource} onClose={() => setSourceOpen(false)} />
      )}
      {tocOptionsOpen && (
        <TocOptionsModal
          editor={editor}
          headingPages={headingPages}
          anchorPage={anchorPage}
          docStyles={docStyles}
          onClose={() => setTocOptionsOpen(false)}
        />
      )}
      {tofOpen && (
        <TofModal
          editor={editor}
          blocks={blocks}
          anchorPage={anchorPage}
          onClose={() => setTofOpen(false)}
        />
      )}
      {xePrompt && (
        <PromptModal
          title={t('ribbonMarkEntryTitle')}
          placeholder={t('ribbonMarkEntryPh')}
          initial={xePrompt.initial}
          onSubmit={submitIndexEntry}
          onClose={() => setXePrompt(null)}
        />
      )}
    </>
  )
}

/* ================= Review ================= */

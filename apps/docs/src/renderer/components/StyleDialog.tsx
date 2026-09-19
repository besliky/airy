/**
 * Word's Modify Style dialog (Home ▸ Styles): edit a paragraph style's
 * definition — name, font, size, bold/italic, color, alignment, spacing and
 * outline level. Only fields the user actually changed are written back to
 * styles.xml through StyleUpsert (BUG-1102: the resolved-display flatten used
 * to pin basedOn-chain values into the style on every modify); a facet the
 * chain supplies is cleared with an explicit off instead of a removal
 * (BUG-1101). The result is pushed live into the document style CSS so every
 * paragraph carrying the pStyle updates on screen at once; the save-side
 * write is surgical (BUG-1001), so untouched facets keep their original bytes,
 * including everything this dialog cannot edit (keepNext, tabs, numPr, …).
 */
import { useState } from 'react'
import type { CSSProperties } from 'react'
import type { Editor } from '@tiptap/core'
import type { DocDefaults, StyleDisplay, StyleInfo, StyleUpsert } from '@airy-office/docx-engine'
import { Dropdown, useModalDialog } from '@airy-office/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import { retagStyleParagraphs } from './ribbon-tabs'

const TWIPS_PER_PT = 20
const ptToTwips = (pt: number) => Math.round(pt * TWIPS_PER_PT)
const twipsToPt = (twips: number | undefined) =>
  twips === undefined ? 0 : Math.round((twips / TWIPS_PER_PT) * 10) / 10

/** styles Word treats as built-in: the upsert must not mark them w:customStyle */
function isBuiltinStyleId(styleId: string): boolean {
  return /^(Normal|Heading[1-9])$/i.test(styleId)
}

export interface StyleEdits {
  name: string
  /** latin face, '' = inherit */
  font: string
  /** east-asian face, '' = inherit (keeps the latin one) */
  fontEa: string
  /** font size in pt, 0 = inherit */
  sizePt: number
  bold: boolean
  italic: boolean
  /** hex without '#', '' = inherit/auto */
  color: string
  align: '' | 'left' | 'center' | 'right' | 'justify'
  beforePt: number
  afterPt: number
  /** line-spacing multiple, 0 = keep a non-auto rule untouched */
  lineSpacing: number
  /** 0 = body text, 1-9 = outline level */
  outline: number
}

/** seed the dialog from the style's resolved definition (docDefaults size fallback) */
export function styleEditsFromInfo(
  info: StyleInfo | undefined,
  docDefaults?: DocDefaults,
): StyleEdits {
  const d = info?.display
  const fontAscii = d?.fontAscii ?? ''
  const fontEa = d?.font && d.font !== fontAscii ? d.font : ''
  return {
    name: info?.name ?? '',
    font: fontAscii,
    fontEa,
    sizePt: Math.round(((d?.sizeHalfPoints ?? docDefaults?.sizeHalfPoints ?? 22) / 2) * 10) / 10,
    bold: d?.bold === true,
    italic: d?.italic === true,
    color: d?.color && d.color !== 'auto' ? d.color : '',
    align: d?.align ?? '',
    beforePt: twipsToPt(d?.spaceBeforeTwips),
    afterPt: twipsToPt(d?.spaceAfterTwips),
    lineSpacing: d?.lineSpacing ?? 0,
    outline: info?.headingLevel ?? 0,
  }
}

/**
 * Flatten the edits into the save-side upsert plus the style's next resolved
 * display (for the live document style CSS) and its new outline level.
 * BUG-1102: only fields the user actually changed are written — an untouched
 * value stays `undefined` so the surgical engine keeps the definition's own
 * bytes and a basedOn parent keeps supplying inherited facets to the child
 * (the resolved-display flatten used to pin chain values into every style it
 * modified, stopping inheritance from propagating). Facets the dialog cannot
 * edit at all (underline, strike, indents, non-auto line rules) are never
 * written here for the same reason.
 *
 * BUG-1101: clearing a facet the basedOn chain supplies writes an explicit off
 * (w:val="0" / w:color "auto" / w:outlineLvl 9) — removing only the style's own
 * element would silently re-inherit the facet after reopen. The live `next`
 * display mirrors the same resolution, so screen and file agree.
 */
export function styleUpsertFromEdits(
  info: StyleInfo | undefined,
  edits: StyleEdits,
  docDefaults?: DocDefaults,
): { upsert: StyleUpsert; display: StyleDisplay | undefined; headingLevel: number | null } {
  const styleId = info?.styleId ?? ''
  const d = info?.display
  // what the chain keeps supplying once the style's own facet is gone
  const chain = info?.chainDisplay
  // the seed the dialog opened with: unchanged fields are not written
  const seed = styleEditsFromInfo(info, docDefaults)
  const hadEaFace = Boolean(d?.font && d.font !== (d?.fontAscii ?? ''))
  const rPr: NonNullable<StyleUpsert['rPr']> = {}
  // value = set, false = explicit off (chain supplies it), null = remove the
  // own element, undefined = untouched (keep the definition's own XML)
  if (edits.bold !== seed.bold)
    rPr.bold = edits.bold
      ? true
      : d?.bold === true
        ? chain?.bold === true
          ? false
          : null
        : undefined
  if (edits.italic !== seed.italic)
    rPr.italic = edits.italic
      ? true
      : d?.italic === true
        ? chain?.italic === true
          ? false
          : null
        : undefined
  // 'auto' is color's explicit-off spelling: the chain would win after a removal
  const colorInChain = chain?.color !== undefined && chain.color !== 'auto'
  if (edits.color !== seed.color)
    rPr.color = edits.color
      ? edits.color.toUpperCase()
      : colorInChain
        ? 'auto'
        : d?.color && d.color !== 'auto'
          ? null
          : undefined
  // 0 = the user cleared the resolved size (docDefaults fallback seed)
  if (edits.sizePt !== seed.sizePt)
    rPr.sizeHalfPoints = edits.sizePt > 0 ? Math.round(edits.sizePt * 2) : null
  if (edits.font !== seed.font) rPr.font = edits.font ? edits.font : d?.fontAscii ? null : undefined
  if (edits.fontEa !== seed.fontEa)
    rPr.fontEa = edits.fontEa ? edits.fontEa : hadEaFace ? null : undefined
  const pPr: NonNullable<StyleUpsert['pPr']> = {}
  // BUG-1002/1102: an interval is written only when the user changed the
  // seeded value — an untouched interval must stay inherited (writing it, even
  // as w:before/after="0", would pin the style over docDefaults or the chain)
  if (edits.beforePt !== seed.beforePt) pPr.spaceBeforeTwips = ptToTwips(edits.beforePt)
  if (edits.afterPt !== seed.afterPt) pPr.spaceAfterTwips = ptToTwips(edits.afterPt)
  if (edits.align !== seed.align)
    pPr.align = edits.align ? edits.align : d?.align ? null : undefined
  // same dirty rule for the spacing multiple: an untouched chain value stays inherited
  if (edits.lineSpacing > 0 && edits.lineSpacing !== seed.lineSpacing)
    pPr.lineSpacing = edits.lineSpacing
  // false = explicit body text (w:outlineLvl 9): a removal would re-inherit
  // the parent heading level after reopen
  if (edits.outline !== seed.outline)
    pPr.outlineLevel =
      edits.outline > 0 ? edits.outline : (info?.headingLevel ?? 0) > 0 ? false : undefined
  const upsert: StyleUpsert = {
    styleId,
    type: 'paragraph',
    name: edits.name.trim() || info?.name || styleId,
    ...(isBuiltinStyleId(styleId) ? { builtin: true } : {}),
    rPr,
    pPr,
  }

  const next: StyleDisplay = { ...(d ?? {}) }
  if (edits.bold) next.bold = true
  // explicit off resolves to false after reopen — keep screen and file alike
  // (rPr.bold false = cleared now; d.bold false = an untouched authored off)
  else if (rPr.bold === false || d?.bold === false) next.bold = false
  else delete next.bold
  if (edits.italic) next.italic = true
  else if (rPr.italic === false || d?.italic === false) next.italic = false
  else delete next.italic
  if (edits.sizePt > 0) next.sizeHalfPoints = Math.round(edits.sizePt * 2)
  // 0 = inherit: the chain keeps supplying its size after the own element clears
  else if (chain?.sizeHalfPoints !== undefined) next.sizeHalfPoints = chain.sizeHalfPoints
  else delete next.sizeHalfPoints
  if (edits.color) next.color = edits.color.toUpperCase()
  else if (rPr.color === 'auto' || d?.color === 'auto') next.color = 'auto'
  else delete next.color
  if (edits.font || edits.fontEa) {
    next.fontAscii = edits.font || undefined
    next.font = edits.fontEa || edits.font || undefined
    if (next.font === undefined) delete next.font
    if (next.fontAscii === undefined) delete next.fontAscii
  } else if (chain?.fontAscii) {
    // inherit: the chain's faces keep applying once the own slots clear
    next.fontAscii = chain.fontAscii
    next.font = chain.font
  } else {
    delete next.font
    delete next.fontAscii
  }
  if (edits.align) next.align = edits.align
  // w:jc has no off spelling: "default" keeps inheriting the chain's alignment
  else if (chain?.align) next.align = chain.align
  else delete next.align
  if (typeof pPr.spaceBeforeTwips === 'number') next.spaceBeforeTwips = pPr.spaceBeforeTwips
  if (typeof pPr.spaceAfterTwips === 'number') next.spaceAfterTwips = pPr.spaceAfterTwips
  if (typeof pPr.lineSpacing === 'number') next.lineSpacing = pPr.lineSpacing
  return {
    upsert,
    display: Object.keys(next).length > 0 ? next : undefined,
    headingLevel: edits.outline > 0 ? edits.outline : null,
  }
}

const ALIGN_KEYS: Array<{ key: StyleEdits['align']; labelKey: StringKey }> = [
  { key: 'left', labelKey: 'appAlignLeft' },
  { key: 'center', labelKey: 'appAlignCenter' },
  { key: 'right', labelKey: 'appAlignRight' },
  { key: 'justify', labelKey: 'appAlignJustify' },
]

interface StyleDialogProps {
  editor: Editor
  /** document style table (ParsedDoc.styles) */
  styles: Map<string, StyleInfo>
  docDefaults?: DocDefaults
  /** paragraph style being modified */
  styleId: string
  onClose: () => void
  /** persist the upsert (save path) and push the live display update */
  onApply: (
    upsert: StyleUpsert,
    display: StyleDisplay | undefined,
    headingLevel: number | null,
  ) => void
}

export function StyleDialog({
  editor,
  styles,
  docDefaults,
  styleId,
  onClose,
  onApply,
}: StyleDialogProps) {
  const { t } = useI18n()
  const dialog = useModalDialog(onClose)
  const info = styles.get(styleId)
  const [edits, setEdits] = useState<StyleEdits>(() => styleEditsFromInfo(info, docDefaults))
  const set = (patch: Partial<StyleEdits>) => setEdits((prev) => ({ ...prev, ...patch }))

  const previewCss: CSSProperties = {
    ...(edits.font || edits.fontEa
      ? {
          fontFamily: `${edits.fontEa || edits.font}${edits.font && edits.fontEa ? `, ${edits.font}` : ''}`,
        }
      : {}),
    ...(edits.sizePt > 0 ? { fontSize: `${edits.sizePt}pt` } : {}),
    ...(edits.bold ? { fontWeight: 'bold' } : {}),
    ...(edits.italic ? { fontStyle: 'italic' } : {}),
    ...(edits.color ? { color: `#${edits.color}` } : {}),
    ...(edits.align === 'center' || edits.align === 'right' || edits.align === 'justify'
      ? { textAlign: edits.align }
      : {}),
  }

  const apply = () => {
    if (!info) {
      onClose()
      return
    }
    const { upsert, display, headingLevel } = styleUpsertFromEdits(info, edits, docDefaults)
    onApply(upsert, display, headingLevel)
    // outline-level change retags the paragraphs already carrying the style
    if (headingLevel !== (info.headingLevel ?? null))
      retagStyleParagraphs(editor, styleId, headingLevel)
    onClose()
  }

  const numInput = (
    label: string,
    value: number,
    pick: (v: number) => void,
    opts?: { min?: number; max?: number; step?: number },
  ) => (
    <label>
      {label}
      <span className="para-num">
        <input
          type="number"
          min={opts?.min ?? 0}
          max={opts?.max ?? 400}
          step={opts?.step ?? 1}
          value={value}
          onChange={(e) => pick(Math.max(opts?.min ?? 0, Number(e.target.value) || 0))}
        />
        <span className="para-unit">{t('ribbonPt')}</span>
      </span>
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('ribbonModifyStyleTitle')}</h2>
        <p className="style-dialog-name">{info?.name ?? styleId}</p>
        <div className="font-dialog-row">
          <label>
            {t('ribbonStyleNameLabel')}
            <input type="text" value={edits.name} onChange={(e) => set({ name: e.target.value })} />
          </label>
          <label>
            {t('ribbonStyleFontLabel')}
            <input type="text" value={edits.font} onChange={(e) => set({ font: e.target.value })} />
          </label>
          <label>
            {t('ribbonStyleFontEaLabel')}
            <input
              type="text"
              value={edits.fontEa}
              onChange={(e) => set({ fontEa: e.target.value })}
            />
          </label>
          {numInput(t('appFontSizeLabel'), edits.sizePt, (v) => set({ sizePt: v }), {
            min: 0,
            max: 400,
            step: 0.5,
          })}
        </div>
        <div className="font-dialog-row">
          <label>
            {t('appFontColor')}
            <span className="style-dialog-color">
              <input
                type="color"
                value={edits.color ? `#${edits.color}` : '#000000'}
                onChange={(e) => set({ color: e.target.value.slice(1).toUpperCase() })}
                aria-label={t('appFontColor')}
              />
              <input
                type="text"
                value={edits.color}
                onChange={(e) =>
                  set({ color: e.target.value.replace(/[^0-9a-fA-F]/g, '').toUpperCase() })
                }
                placeholder="auto"
              />
            </span>
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={edits.bold}
              onChange={(e) => set({ bold: e.target.checked })}
            />
            {t('appFontBold')}
          </label>
          <label className="font-check">
            <input
              type="checkbox"
              checked={edits.italic}
              onChange={(e) => set({ italic: e.target.checked })}
            />
            {t('appFontItalic')}
          </label>
        </div>
        <div className="font-dialog-row">
          <label>
            {t('appAlignment')}
            <Dropdown
              value={edits.align || 'default'}
              ariaLabel={t('appAlignment')}
              options={[
                { value: 'default', label: t('ribbonStyleAlignDefault') },
                ...ALIGN_KEYS.map((o) => ({ value: o.key as string, label: t(o.labelKey) })),
              ]}
              onPick={(v) => set({ align: v === 'default' ? '' : (v as StyleEdits['align']) })}
            />
          </label>
          {numInput(t('appSpaceBefore'), edits.beforePt, (v) => set({ beforePt: v }))}
          {numInput(t('appSpaceAfter'), edits.afterPt, (v) => set({ afterPt: v }))}
          <label>
            {t('appLineSpacingLabel')}
            <span className="para-num">
              <input
                type="number"
                min={0}
                max={10}
                step={0.05}
                value={edits.lineSpacing}
                onChange={(e) => set({ lineSpacing: Math.max(0, Number(e.target.value) || 0) })}
              />
              <span className="para-unit">×</span>
            </span>
          </label>
        </div>
        <div className="font-dialog-row">
          <label>
            {t('ribbonStyleOutlineLevel')}
            <Dropdown
              value={String(edits.outline)}
              ariaLabel={t('ribbonStyleOutlineLevel')}
              options={[
                { value: '0', label: t('ribbonStyleOutlineBody') },
                ...Array.from({ length: 9 }, (_, i) => ({
                  value: String(i + 1),
                  label: t('ribbonStyleOutlineLevelN', { n: i + 1 }),
                })),
              ]}
              onPick={(v) => set({ outline: Number(v) })}
            />
          </label>
          <span className="style-dialog-preview" style={previewCss}>
            {t('appFontPreviewSample')}
          </span>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply} disabled={!info}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

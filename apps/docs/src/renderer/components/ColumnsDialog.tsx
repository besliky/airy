import { useState } from 'react'
import type { SectionSettings } from '@airy-office/docx-engine'
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from '../i18n/locale'

const TWIPS_PER_CM = 567

/** Word's Left/Right two-column presets split the text area ~1.69:1 */
export const UNEQUAL_COLUMN_RATIO = 1.69

export type ColumnPreset = 'one' | 'two' | 'three' | 'left' | 'right'

const PRESET_COLUMNS: Record<ColumnPreset, number> = {
  one: 1,
  two: 2,
  three: 3,
  left: 2,
  right: 2,
}

/**
 * Explicit column widths (twips) of a preset over a text area: equal presets
 * divide the remainder evenly (undefined = equalWidth), Left/Right split it
 * 1.69:1 (narrow column last for Left, first for Right).
 */
export function columnPresetSpec(
  preset: ColumnPreset,
  contentWidthTwips: number,
  spacingTwips: number,
): { columns: number; colWidths?: number[] } {
  const columns = PRESET_COLUMNS[preset]
  if (preset === 'one' || preset === 'two' || preset === 'three') {
    return { columns }
  }
  const total = Math.max(spacingTwips, contentWidthTwips - spacingTwips)
  const wide = Math.round((total * UNEQUAL_COLUMN_RATIO) / (UNEQUAL_COLUMN_RATIO + 1))
  const narrow = Math.max(567, contentWidthTwips - spacingTwips - wide)
  return { columns, colWidths: preset === 'left' ? [wide, narrow] : [narrow, wide] }
}

/** which preset a section's current column setup matches (equal widths win) */
export function presetOf(section: SectionSettings): ColumnPreset {
  if (section.columns <= 1) return 'one'
  if (section.columns >= 3) return 'three'
  const w = section.colWidths
  if (!w || w.length !== 2) return 'two'
  const [a, b] = w
  if (Math.abs(a - b) <= 56) return 'two'
  return a > b ? 'left' : 'right'
}

const cm = (twips: number) => Math.round((twips / TWIPS_PER_CM) * 100) / 100

const num = (raw: string, min: number, max: number, fallback: number): number => {
  const v = Number(raw)
  if (!Number.isFinite(v)) return fallback
  return Math.min(max, Math.max(min, Math.round(v * TWIPS_PER_CM)))
}

/**
 * Word's Layout → Columns dialog: One/Two/Three + Left/Right presets, column
 * spacing, the first column's width for the unequal presets, and the vertical
 * separator line between columns (w:cols w:sep).
 */
export function ColumnsDialog({
  section,
  onApply,
  onClose,
}: {
  section: SectionSettings
  onApply: (next: SectionSettings) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const contentW = section.pageWidth - section.marginLeft - section.marginRight
  const [preset, setPreset] = useState<ColumnPreset>(presetOf(section))
  const [spacing, setSpacing] = useState(String(cm(section.colSpace ?? 720)))
  const unequal = preset === 'left' || preset === 'right'
  const currentW = section.colWidths?.[0]
  const [width, setWidth] = useState(
    String(
      cm(
        currentW ??
          Math.round(
            (contentW - (section.colSpace ?? 720)) *
              (unequal ? UNEQUAL_COLUMN_RATIO / (UNEQUAL_COLUMN_RATIO + 1) : 0.5),
          ),
      ),
    ),
  )
  const [lineBetween, setLineBetween] = useState(section.columnSep === true)
  const dialog = useModalDialog(onClose)

  const submit = () => {
    const space = num(spacing, 0, 2834, 720)
    let spec = columnPresetSpec(preset, contentW, space)
    if (unequal) {
      const first = Math.min(
        contentW - space - 567,
        Math.max(567, num(width, 1, 100, contentW / 2)),
      )
      spec = { columns: 2, colWidths: [first, Math.max(567, contentW - space - first)] }
    }
    const next: SectionSettings = { ...section, columns: spec.columns, colSpace: space }
    if (spec.colWidths) next.colWidths = spec.colWidths
    else delete next.colWidths
    if (lineBetween) next.columnSep = true
    else delete next.columnSep
    onApply(next)
    onClose()
  }

  const presetButton = (
    key: ColumnPreset,
    labelKey:
      'layoutColOne' | 'layoutColTwo' | 'layoutColThree' | 'layoutColLeft' | 'layoutColRight',
  ) => (
    <button
      key={key}
      className={`cols-preset ${preset === key ? 'active' : ''}`}
      aria-pressed={preset === key}
      onClick={() => setPreset(key)}
    >
      {t(labelKey)}
    </button>
  )

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('layoutColsDialogTitle')}</h2>
        <div className="modal-row">
          {presetButton('one', 'layoutColOne')}
          {presetButton('two', 'layoutColTwo')}
          {presetButton('three', 'layoutColThree')}
          {presetButton('left', 'layoutColLeft')}
          {presetButton('right', 'layoutColRight')}
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('layoutColSpacing')}
            <input
              type="number"
              min={0}
              max={5}
              step={0.05}
              disabled={preset === 'one'}
              value={spacing}
              onChange={(e) => setSpacing(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && submit()}
            />
          </label>
          <label>
            {unequal ? t('layoutColWidth1') : t('layoutColWidth')}
            <input
              type="number"
              min={1}
              max={50}
              step={0.05}
              disabled={!unequal}
              value={width}
              onChange={(e) => setWidth(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && submit()}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label className="ln-auto">
            <input
              type="checkbox"
              checked={lineBetween}
              onChange={(e) => setLineBetween(e.target.checked)}
            />
            {t('layoutLineBetween')}
          </label>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

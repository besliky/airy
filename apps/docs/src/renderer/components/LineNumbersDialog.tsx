import { useState } from 'react'
import { useModalDialog } from '@airy-office/ui'
import type { LineNumberSettings } from '@airy-office/docx-engine'
import { useI18n } from '../i18n/locale'

const TWIPS_PER_PT = 20

/** clamp + commit-on-blur number input (matches the MarginDialog field style) */
const num = (raw: string, min: number, max: number, fallback: number): number => {
  const v = Number(raw)
  if (!Number.isFinite(v)) return fallback
  return Math.min(max, Math.max(min, Math.round(v)))
}

/**
 * Word's Layout → Line Numbers options dialog: restart mode, start value,
 * count increment and the numeral's distance from the text. Apply yields
 * undefined for "None" (numbering off).
 */
export function LineNumbersDialog({
  value,
  onApply,
  onClose,
}: {
  value: LineNumberSettings | undefined
  onApply: (next: LineNumberSettings | undefined) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [restart, setRestart] = useState<'none' | NonNullable<LineNumberSettings['restart']>>(
    value?.restart ?? 'newPage',
  )
  const [startAt, setStartAt] = useState(String(value?.start ?? 1))
  const [countBy, setCountBy] = useState(String(value?.countBy ?? 1))
  const [distAuto, setDistAuto] = useState(value?.distance === undefined)
  const [distance, setDistance] = useState(
    String(value?.distance !== undefined ? value.distance / TWIPS_PER_PT : 24),
  )
  const dialog = useModalDialog(onClose)

  const submit = () => {
    if (restart === 'none') {
      onApply(undefined)
    } else {
      onApply({
        restart,
        start: num(startAt, 0, 9999, 1),
        countBy: num(countBy, 1, 100, 1),
        ...(distAuto ? {} : { distance: num(distance, 0, 3168, 24) * TWIPS_PER_PT }),
      })
    }
    onClose()
  }

  const radio = (
    mode: 'none' | NonNullable<LineNumberSettings['restart']>,
    key:
      | 'ribbonLineNumbersNone'
      | 'ribbonLineNumbersRestartPage'
      | 'ribbonLineNumbersRestartSection'
      | 'ribbonLineNumbersContinuous',
  ) => (
    <label key={mode} className="ln-radio">
      <input
        type="radio"
        name="ln-restart"
        checked={restart === mode}
        onChange={() => setRestart(mode)}
      />
      {t(key)}
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('ribbonLnDialogTitle')}</h2>
        <div className="modal-row ln-modes">
          {radio('none', 'ribbonLineNumbersNone')}
          {radio('newPage', 'ribbonLineNumbersRestartPage')}
          {radio('newSection', 'ribbonLineNumbersRestartSection')}
          {radio('continuous', 'ribbonLineNumbersContinuous')}
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonLnStartAt')}
            <input
              type="number"
              min={0}
              max={9999}
              step={1}
              value={startAt}
              onChange={(e) => setStartAt(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          </label>
          <label>
            {t('ribbonLnCountBy')}
            <input
              type="number"
              min={1}
              max={100}
              step={1}
              value={countBy}
              onChange={(e) => setCountBy(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonLnDistance')}
            <input
              type="number"
              min={0}
              max={3168}
              step={1}
              value={distance}
              disabled={distAuto}
              onChange={(e) => setDistance(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          </label>
          <label className="ln-auto">
            <input
              type="checkbox"
              checked={distAuto}
              onChange={(e) => setDistAuto(e.target.checked)}
            />
            {t('ribbonLnDistanceAuto')}
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

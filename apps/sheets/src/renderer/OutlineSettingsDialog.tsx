import { useState } from 'react'
import { useI18n } from './i18n/locale'

/// Data → Outline → Settings: the sheet's summary placement (sheetPr
/// outlinePr). The placement drives where summary lines — and the gutter's
/// +/- buttons — sit relative to each group's detail span.

export interface OutlineSettingsValue {
  readonly summaryBelow: boolean
  readonly summaryRight: boolean
}

export function OutlineSettingsDialog({
  initial,
  onApply,
  onClose,
}: {
  readonly initial: OutlineSettingsValue
  /// Returns an error message, or null on success.
  readonly onApply: (value: OutlineSettingsValue) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [summaryBelow, setSummaryBelow] = useState(initial.summaryBelow)
  const [summaryRight, setSummaryRight] = useState(initial.summaryRight)
  const [error, setError] = useState<string | null>(null)

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgOutlineSettingsTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgOutlineSettingsTitle')}</header>
        <div className="dialog-grid">
          <fieldset className="t2c-fieldset">
            <legend>{t('dlgOutlineRows')}</legend>
            <label className="print-radio">
              <input
                type="radio"
                name="outline-summary-rows"
                checked={summaryBelow}
                onChange={() => setSummaryBelow(true)}
              />
              {t('dlgOutlineSummaryBelow')}
            </label>
            <label className="print-radio">
              <input
                type="radio"
                name="outline-summary-rows"
                checked={!summaryBelow}
                onChange={() => setSummaryBelow(false)}
              />
              {t('dlgOutlineSummaryAbove')}
            </label>
          </fieldset>
          <fieldset className="t2c-fieldset">
            <legend>{t('dlgOutlineColumns')}</legend>
            <label className="print-radio">
              <input
                type="radio"
                name="outline-summary-cols"
                checked={summaryRight}
                onChange={() => setSummaryRight(true)}
              />
              {t('dlgOutlineSummaryRight')}
            </label>
            <label className="print-radio">
              <input
                type="radio"
                name="outline-summary-cols"
                checked={!summaryRight}
                onChange={() => setSummaryRight(false)}
              />
              {t('dlgOutlineSummaryLeft')}
            </label>
          </fieldset>
          <p className="dialog-note">{t('dlgOutlineSettingsNote')}</p>
        </div>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            onClick={() => {
              const failure = onApply({ summaryBelow, summaryRight })
              setError(failure)
              if (failure === null) onClose()
            }}
          >
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

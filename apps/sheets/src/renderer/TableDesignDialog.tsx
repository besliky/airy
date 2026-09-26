import { useState } from 'react'
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from './i18n/locale'
import { areaToRangeText, parseRangeText, type TableDesignSeed } from './table-design'

/// Table Design (PAR-202): rename, resize (header-anchored), style and
/// banding, and Convert to Range for the table under the active cell.
/// Applying routes through the App-level callback so the journal and the
/// status bar stay in one place; the dialog shows the returned error.

const TABLE_STYLES = [
  'TableStyleMedium2',
  'TableStyleMedium4',
  'TableStyleMedium7',
  'TableStyleLight1',
  'TableStyleLight9',
  'TableStyleDark2',
] as const

export function TableDesignDialog({
  seed,
  onApply,
  onClose,
}: {
  readonly seed: TableDesignSeed
  /// Returns an error message, or null on success. `converted` tells the
  /// caller the table became a range (different status message).
  readonly onApply: (change: {
    readonly seed: TableDesignSeed
    readonly name: string
    readonly areaText: string
    readonly style: string | undefined
    readonly bandedRows: boolean
    readonly convertToRange: boolean
  }) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const dialog = useModalDialog(onClose)
  const [name, setName] = useState(seed.name)
  const [range, setRange] = useState(areaToRangeText(seed.area))
  const [style, setStyle] = useState<string>(seed.style ?? 'TableStyleMedium2')
  const [bandedRows, setBandedRows] = useState(seed.bandedRows)
  const [confirming, setConfirming] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const apply = (convertToRange: boolean): void => {
    if (convertToRange && !confirming) {
      setConfirming(true)
      return
    }
    const failure = onApply({
      seed,
      name: name.trim(),
      areaText: range.trim(),
      style,
      bandedRows,
      convertToRange,
    })
    setError(failure)
    if (failure === null) onClose()
  }

  return (
    <div className="dialog-backdrop" {...dialog.backdropProps} onClick={onClose}>
      <div
        className="format-cells-dialog"
        {...dialog.dialogProps}
        onClick={(event) => event.stopPropagation()}
      >
        <header {...dialog.titleProps}>{t('dlgTableTitle')}</header>
        <div className="dialog-grid">
          <label className="print-radio dialog-field">
            <span>{t('dlgTableName')}</span>
            <input
              type="text"
              value={name}
              onChange={(event) => {
                setName(event.target.value)
                setConfirming(false)
              }}
              autoFocus
            />
          </label>
          <label className="print-radio dialog-field">
            <span>{t('dlgTableRange')}</span>
            <input
              type="text"
              value={range}
              onChange={(event) => {
                setRange(event.target.value)
                setConfirming(false)
              }}
            />
          </label>
          <label className="print-radio dialog-field">
            <span>{t('dlgTableStyle')}</span>
            <select value={style} onChange={(event) => setStyle(event.target.value)}>
              {TABLE_STYLES.map((styleName) => (
                <option key={styleName} value={styleName}>
                  {styleName}
                </option>
              ))}
            </select>
          </label>
          <label className="print-radio">
            <input
              type="checkbox"
              checked={bandedRows}
              onChange={(event) => setBandedRows(event.target.checked)}
            />
            {t('dlgTableBandedRows')}
          </label>
          <p className="dialog-note">{t('dlgTableNote')}</p>
        </div>
        {confirming && (
          <p className="dialog-note" role="alert">
            {t('dlgTableConvertConfirm')}
          </p>
        )}
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="secondary" onClick={() => apply(true)}>
            {t('dlgTableConvert')}
          </button>
          <button className="primary-action" onClick={() => apply(false)}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

// Re-exported so ExcelShell/App need only this module for the dialog.
export { areaToRangeText, parseRangeText }

import { useEffect, useMemo, useState } from 'react'
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from './i18n/locale'
import {
  activeDelimiterChars,
  coerceFieldValue,
  DEFAULT_DELIMITERS,
  parseBreakPositions,
  splitDelimited,
  splitFixedWidth,
  type TextToColumnsConfig,
  type TextToColumnsDelimiters,
  type TextToColumnsMode,
  type TextToColumnType,
} from './text-to-columns'

/// Excel's Data → Text to Columns wizard, condensed into one pane: the
/// split mode (delimited / fixed width), the delimiter set or break
/// positions, a data-type select per output column, and the destination
/// cell. The preview shows the parsed result of the first rows.

const PREVIEW_ROWS = 12

export interface TextToColumnsSource {
  readonly rows: readonly string[]
  readonly destinationLabel: string
}

export function TextToColumnsDialog({
  source,
  onApply,
  onDestinationOverwrites,
  onClose,
}: {
  readonly source: TextToColumnsSource
  /// Returns an error message, or null on success.
  readonly onApply: (config: TextToColumnsConfig) => string | null
  /// Whether the config's destination rectangle would overwrite non-empty
  /// cells outside the split column — the wizard asks before that apply. On
  /// streamed workbooks the probe reads the file floor through the sidecar
  /// (BUG-1312), so the answer can arrive as a promise.
  readonly onDestinationOverwrites?: (config: TextToColumnsConfig) => boolean | Promise<boolean>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const dialog = useModalDialog(onClose)
  const [mode, setMode] = useState<TextToColumnsMode>('delimited')
  const [delimiters, setDelimiters] = useState<TextToColumnsDelimiters>(DEFAULT_DELIMITERS)
  const [breaksText, setBreaksText] = useState('')
  const [columnTypes, setColumnTypes] = useState<TextToColumnType[]>(['general'])
  const [destination, setDestination] = useState('')
  const [error, setError] = useState<string | null>(null)
  /// Armed by the first OK/Enter when the destination would overwrite
  /// unrelated data; the second one applies. Any config edit disarms it.
  const [confirmOverwrite, setConfirmOverwrite] = useState(false)
  /// True while an async overwrite probe is in flight — a second OK/Enter
  /// must not start a competing apply before the probe has answered.
  const [probing, setProbing] = useState(false)

  useEffect(() => {
    setConfirmOverwrite(false)
  }, [mode, delimiters, breaksText, columnTypes, destination])

  const breaks = useMemo(() => parseBreakPositions(breaksText), [breaksText])
  const activeDelimiters = useMemo(() => activeDelimiterChars(delimiters), [delimiters])

  const preview = useMemo(() => {
    const split = (text: string): string[] =>
      mode === 'fixed-width'
        ? splitFixedWidth(text, breaks ?? [])
        : activeDelimiters.length === 0
          ? [text]
          : splitDelimited(text, activeDelimiters, delimiters.consecutiveAsOne)
    return source.rows.slice(0, PREVIEW_ROWS).map(split)
  }, [mode, breaks, activeDelimiters, delimiters.consecutiveAsOne, source.rows])

  const width = preview.reduce((max, row) => Math.max(max, row.length), 0)
  const typeAt = (column: number): TextToColumnType => columnTypes[column] ?? 'general'

  const modeError =
    mode === 'delimited' && activeDelimiters.length === 0
      ? t('dlgT2cNeedDelimiter')
      : mode === 'fixed-width' && breaksText.trim() !== '' && breaks === null
        ? t('dlgT2cBadBreaks')
        : null

  const commitApply = (config: TextToColumnsConfig): void => {
    const failure = onApply(config)
    setError(failure)
    if (failure !== null) {
      // A failed apply disarms the confirm: the retry re-runs the overwrite
      // probe instead of sailing through on the stale go-ahead.
      setConfirmOverwrite(false)
      return
    }
    onClose()
  }

  const apply = (): void => {
    if (modeError !== null || width === 0 || probing) return
    const config: TextToColumnsConfig = {
      mode,
      delimiters,
      breaks: breaks ?? [],
      columnTypes: Array.from({ length: width }, (_, column) => typeAt(column)),
      destination: destination.trim() === '' ? null : destination.trim(),
    }
    // Excel asks before the fields replace unrelated destination cells: the
    // first OK/Enter arms a visible confirm instead of applying, the second
    // one goes through (a config edit disarms it again).
    const asks = !confirmOverwrite ? onDestinationOverwrites?.(config) : undefined
    if (asks !== undefined && typeof asks !== 'boolean') {
      // Streamed workbooks probe through the file floor (a sidecar read), so
      // the answer is asynchronous; double commits stay blocked meanwhile.
      setProbing(true)
      void asks
        .then((overwrites) => {
          if (overwrites) {
            setConfirmOverwrite(true)
            return
          }
          commitApply(config)
        })
        .finally(() => setProbing(false))
      return
    }
    if (asks === true) {
      setConfirmOverwrite(true)
      return
    }
    commitApply(config)
  }

  /// Enter in any of the wizard's text fields commits, like the neighboring
  /// sheets one-liner dialogs (the modal hook owns Escape, not Enter).
  const commitOnEnter = (event: React.KeyboardEvent<HTMLInputElement>): void => {
    if (event.key !== 'Enter') return
    event.preventDefault()
    apply()
  }

  const checkbox = (
    key: keyof Omit<TextToColumnsDelimiters, 'custom' | 'consecutiveAsOne'>,
    label: string,
  ): React.JSX.Element => (
    <label className="print-radio">
      <input
        type="checkbox"
        checked={delimiters[key]}
        onChange={(event) => setDelimiters({ ...delimiters, [key]: event.target.checked })}
      />
      {label}
    </label>
  )

  return (
    <div className="dialog-backdrop" {...dialog.backdropProps} onClick={onClose}>
      <div
        className="format-cells-dialog"
        {...dialog.dialogProps}
        onClick={(event) => event.stopPropagation()}
      >
        <header {...dialog.titleProps}>{t('dlgT2cTitle')}</header>
        <div className="dialog-grid">
          <fieldset className="t2c-fieldset">
            <legend>{t('dlgT2cMode')}</legend>
            <label className="print-radio">
              <input
                type="radio"
                name="t2c-mode"
                checked={mode === 'delimited'}
                onChange={() => setMode('delimited')}
              />
              {t('dlgT2cDelimited')}
            </label>
            <label className="print-radio">
              <input
                type="radio"
                name="t2c-mode"
                checked={mode === 'fixed-width'}
                onChange={() => setMode('fixed-width')}
              />
              {t('dlgT2cFixedWidth')}
            </label>
          </fieldset>
          {mode === 'delimited' ? (
            <fieldset className="t2c-fieldset">
              <legend>{t('dlgT2cDelimiters')}</legend>
              {checkbox('tab', t('dlgT2cTab'))}
              {checkbox('semicolon', t('dlgT2cSemicolon'))}
              {checkbox('comma', t('dlgT2cComma'))}
              {checkbox('space', t('dlgT2cSpace'))}
              <label className="print-radio">
                {t('dlgT2cOther')}
                <input
                  type="text"
                  className="print-range-input t2c-other-input"
                  value={delimiters.custom}
                  onChange={(event) => setDelimiters({ ...delimiters, custom: event.target.value })}
                  onKeyDown={commitOnEnter}
                />
              </label>
              <label className="print-radio">
                <input
                  type="checkbox"
                  checked={delimiters.consecutiveAsOne}
                  onChange={(event) =>
                    setDelimiters({ ...delimiters, consecutiveAsOne: event.target.checked })
                  }
                />
                {t('dlgT2cConsecutive')}
              </label>
            </fieldset>
          ) : (
            <fieldset className="t2c-fieldset">
              <legend>{t('dlgT2cBreakPositions')}</legend>
              <label className="print-radio">
                {t('dlgT2cBreaksAt')}
                <input
                  type="text"
                  className="print-range-input t2c-breaks-input"
                  placeholder="5, 12, 20"
                  value={breaksText}
                  onChange={(event) => setBreaksText(event.target.value)}
                  onKeyDown={commitOnEnter}
                />
              </label>
              <p className="dialog-note">{t('dlgT2cBreaksNote')}</p>
            </fieldset>
          )}
          <label>
            {t('dlgT2cDestination')}
            <input
              type="text"
              className="print-range-input t2c-destination-input"
              placeholder={source.destinationLabel}
              value={destination}
              onChange={(event) => setDestination(event.target.value)}
              onKeyDown={commitOnEnter}
            />
          </label>
        </div>
        {width > 0 && (
          <div className="t2c-preview" aria-label={t('dlgT2cPreview')}>
            <table>
              <thead>
                <tr>
                  {Array.from({ length: width }, (_, column) => (
                    <th key={column}>
                      <select
                        aria-label={t('dlgT2cColumnType')}
                        value={typeAt(column)}
                        onChange={(event) => {
                          const next = [...columnTypes]
                          next[column] = event.target.value as TextToColumnType
                          setColumnTypes(next)
                        }}
                      >
                        <option value="general">{t('dlgT2cTypeGeneral')}</option>
                        <option value="text">{t('dlgT2cTypeText')}</option>
                        <option value="date-dmy">{t('dlgT2cTypeDateDmy')}</option>
                        <option value="date-mdy">{t('dlgT2cTypeDateMdy')}</option>
                        <option value="date-ymd">{t('dlgT2cTypeDateYmd')}</option>
                      </select>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((row, index) => (
                  <tr key={index}>
                    {Array.from({ length: width }, (_, column) => {
                      const text = row[column] ?? ''
                      const type = typeAt(column)
                      const value = coerceFieldValue(text, type)
                      return <td key={column}>{String(value.v ?? '')}</td>
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {confirmOverwrite && (
          <p className="dialog-note" role="alert">
            {t('dlgT2cOverwriteNote')}
          </p>
        )}
        {(error ?? modeError) && (
          <p className="dialog-note" role="alert">
            {error ?? modeError}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            disabled={modeError !== null || width === 0 || probing}
            onClick={apply}
          >
            {confirmOverwrite ? t('dlgT2cOverwriteConfirm') : t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

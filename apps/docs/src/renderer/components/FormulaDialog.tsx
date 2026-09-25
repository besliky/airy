import { useState } from 'react'
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from '../i18n/locale'

/** common Word numeric pictures offered in the Number format combo */
const NUMBER_FORMATS = [
  '',
  '#,##0.00',
  '0',
  '0.00',
  '#,##0',
  '0%',
  '$#,##0.00',
  '#,##0.00;(#,##0.00)',
]

/**
 * Word's Table Layout → Formula dialog: the field instruction (prefilled with
 * =SUM(ABOVE)/=SUM(LEFT) per Word's own heuristic) and the optional numeric
 * picture switch (\# "…"). OK inserts the field with its computed cache.
 */
export function FormulaDialog({
  initialFormula,
  onApply,
  onClose,
}: {
  initialFormula: string
  onApply: (instruction: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [formula, setFormula] = useState(initialFormula)
  const [format, setFormat] = useState('')
  const dialog = useModalDialog(onClose)

  const submit = () => {
    const expr = formula.trim()
    if (!expr.startsWith('=')) return
    onApply(format.trim() ? `${expr} \\# "${format.trim()}"` : expr)
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('ribbonFormulaDialogTitle')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonFormulaField')}
            <input
              type="text"
              value={formula}
              spellCheck={false}
              onChange={(e) => setFormula(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
              // select the expression body (without the leading =) like Word,
              // so typing replaces just the function part
              onFocus={(e) => e.target.setSelectionRange(1, e.target.value.length)}
              autoFocus
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonFormulaNumberFormat')}
            <input
              type="text"
              list="table-formula-formats"
              value={format}
              spellCheck={false}
              placeholder={t('ribbonFormulaNumberFormatNone')}
              onChange={(e) => setFormat(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          </label>
          <datalist id="table-formula-formats">
            {NUMBER_FORMATS.filter((f) => f !== '').map((f) => (
              <option key={f} value={f} />
            ))}
          </datalist>
        </div>
        <p className="modal-hint">{t('ribbonFormulaHint')}</p>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button
            className="btn-primary"
            onClick={submit}
            disabled={!formula.trim().startsWith('=')}
          >
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

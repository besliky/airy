/**
 * Review > Protect Sheet / Unprotect Sheet (PAR-204). Excel's dialog, minus
 * the object/scenario checkboxes this save always writes as locked:
 *
 *  - protect: optional password (typed twice, Excel's confirmation rule) and
 *    the "Allow all users of this worksheet to:" checkbox list. Checkboxes
 *    speak "allowed"; onApply receives the raw OOXML polarity (true =
 *    prevented) the journal and serializer store.
 *  - unprotect: the single password prompt; a wrong password keeps the dialog
 *    open with Excel's message, and a modern hashValue sheet fails closed.
 */
import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from '@airy-office/ui'
import type { SheetProtectionAttributes } from '../shared/desktop-api'

/// The dialog's checkbox list in "allowed" terms with Excel's default check
/// states: selecting cells stays allowed, everything else starts locked.
const ALLOWED_CHECKBOXES: readonly {
  readonly attribute: keyof SheetProtectionAttributes
  readonly labelKey:
    | 'appProtectAllowSelectLocked'
    | 'appProtectAllowSelectUnlocked'
    | 'appProtectAllowFormatCells'
    | 'appProtectAllowFormatColumns'
    | 'appProtectAllowFormatRows'
    | 'appProtectAllowInsertRows'
    | 'appProtectAllowInsertColumns'
    | 'appProtectAllowDeleteRows'
    | 'appProtectAllowDeleteColumns'
    | 'appProtectAllowSort'
    | 'appProtectAllowAutoFilter'
  readonly defaultAllowed: boolean
}[] = [
  { attribute: 'selectLockedCells', labelKey: 'appProtectAllowSelectLocked', defaultAllowed: true },
  {
    attribute: 'selectUnlockedCells',
    labelKey: 'appProtectAllowSelectUnlocked',
    defaultAllowed: true,
  },
  { attribute: 'formatCells', labelKey: 'appProtectAllowFormatCells', defaultAllowed: false },
  { attribute: 'formatColumns', labelKey: 'appProtectAllowFormatColumns', defaultAllowed: false },
  { attribute: 'formatRows', labelKey: 'appProtectAllowFormatRows', defaultAllowed: false },
  { attribute: 'insertRows', labelKey: 'appProtectAllowInsertRows', defaultAllowed: false },
  { attribute: 'insertColumns', labelKey: 'appProtectAllowInsertColumns', defaultAllowed: false },
  { attribute: 'deleteColumns', labelKey: 'appProtectAllowDeleteColumns', defaultAllowed: false },
  { attribute: 'deleteRows', labelKey: 'appProtectAllowDeleteRows', defaultAllowed: false },
  { attribute: 'sort', labelKey: 'appProtectAllowSort', defaultAllowed: false },
  { attribute: 'autoFilter', labelKey: 'appProtectAllowAutoFilter', defaultAllowed: false },
]

export interface ProtectSheetRequest {
  readonly mode: 'protect' | 'unprotect'
  readonly password: string
  /// Raw OOXML polarity (true = prevented); protect mode only.
  readonly attributes: SheetProtectionAttributes
}

export function ProtectSheetDialog({
  mode,
  onApply,
  onClose,
}: {
  readonly mode: 'protect' | 'unprotect'
  /// Returns an error message, or null on success.
  readonly onApply: (request: ProtectSheetRequest) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const dialog = useModalDialog(onClose)
  const [password, setPassword] = useState('')
  const [confirm, setConfirm] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [allowed, setAllowed] = useState<Record<string, boolean>>(() =>
    Object.fromEntries(ALLOWED_CHECKBOXES.map((entry) => [entry.attribute, entry.defaultAllowed])),
  )

  const apply = (): void => {
    if (mode === 'protect' && password !== confirm) {
      setError(t('appProtectPasswordMismatch'))
      return
    }
    const failed = onApply({
      mode,
      password,
      attributes: Object.fromEntries(
        ALLOWED_CHECKBOXES.map((entry) => [entry.attribute, !allowed[entry.attribute]]),
      ),
    })
    if (failed !== null) {
      setError(failed)
      return
    }
    onClose()
  }

  return (
    <div className="dialog-backdrop" {...dialog.backdropProps} onClick={onClose}>
      <div
        className="format-cells-dialog"
        {...dialog.dialogProps}
        onClick={(event) => event.stopPropagation()}
      >
        <header {...dialog.titleProps}>
          {mode === 'protect' ? t('appProtectSheet') : t('appUnprotectSheet')}
        </header>
        <section className="dialog-body">
          {mode === 'protect' ? (
            <>
              <div className="dialog-grid">
                <label>
                  {t('appProtectPassword')}
                  <input
                    autoFocus
                    type="password"
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                  />
                </label>
                <label>
                  {t('appProtectReenterPassword')}
                  <input
                    type="password"
                    value={confirm}
                    onChange={(event) => setConfirm(event.target.value)}
                  />
                </label>
              </div>
              <p className="dialog-note">{t('appProtectAllowedHeading')}</p>
              <div className="protect-sheet-checkboxes">
                {ALLOWED_CHECKBOXES.map((entry) => (
                  <label key={entry.attribute} className="protect-sheet-checkbox">
                    <input
                      type="checkbox"
                      checked={allowed[entry.attribute] ?? false}
                      onChange={(event) =>
                        setAllowed((current) => ({
                          ...current,
                          [entry.attribute]: event.target.checked,
                        }))
                      }
                    />
                    {t(entry.labelKey)}
                  </label>
                ))}
              </div>
              <p className="dialog-note">{t('appProtectionWillWrite')}</p>
            </>
          ) : (
            <>
              <div className="dialog-grid">
                <label>
                  {t('appProtectPassword')}
                  <input
                    autoFocus
                    type="password"
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === 'Enter') apply()
                    }}
                  />
                </label>
              </div>
              <p className="dialog-note">{t('appProtectionWillRemove')}</p>
            </>
          )}
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

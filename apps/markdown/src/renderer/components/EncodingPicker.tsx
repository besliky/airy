import { Dropdown } from '@airy-office/ui'
import { useI18n } from '../i18n/locale'
import { SELECTABLE_ENCODINGS } from '../../shared/ipc'
import type { SelectableEncoding } from '../../shared/ipc'

/** the status-bar encoding value: no override ('auto') or a forced charset (UX-1696) */
export type EncodingPick = 'auto' | SelectableEncoding

/**
 * Status-bar "Reopen with encoding" control (UX-1696): shows the active
 * override ('auto' until a pick is made in this session) and offers
 * auto-detection plus every selectable charset. Picking hands the choice to
 * the caller, which remembers it through setEncoding and re-reads the file.
 */
export function EncodingPicker({
  pick,
  disabled,
  onPick,
}: {
  pick: EncodingPick
  disabled?: boolean
  onPick: (pick: EncodingPick) => void
}) {
  const { t } = useI18n()
  return (
    <Dropdown
      className="encoding-dd"
      value={pick}
      disabled={disabled}
      ariaLabel={t('reopenEncoding')}
      tip={t('reopenEncoding')}
      options={[
        { value: 'auto', label: t('encodingAuto') },
        // charset names are technical identifiers (VS Code convention) — not translated
        ...SELECTABLE_ENCODINGS.map((encoding) => ({ value: encoding, label: encoding })),
      ]}
      onPick={onPick}
    />
  )
}

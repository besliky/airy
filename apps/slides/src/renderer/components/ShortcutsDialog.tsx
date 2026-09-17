/**
 * PowerPoint's keyboard-shortcut reference (Help ▸ Keyboard Shortcuts).
 * Read-only: it renders the registry in shortcut-registry.ts, so a binding
 * added there shows up without touching this file. Same pattern as the docs
 * app's ShortcutsDialog.
 */
import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { SHORTCUT_GROUPS, SHORTCUTS, shortcutKeys } from '../shortcut-registry'
import { useModalKeys } from './modal-keys'

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [query, setQuery] = useState('')

  const needle = query.trim().toLowerCase()
  const rows = SHORTCUTS.map((def) => ({
    id: def.id,
    group: def.group,
    label: t(def.labelKey),
    keys: shortcutKeys(def),
  })).filter(
    (row) =>
      !needle ||
      row.label.toLowerCase().includes(needle) ||
      row.keys.toLowerCase().includes(needle),
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal modal-shortcuts" role="dialog" aria-label={t('scTitle')}>
        <h2>{t('scTitle')}</h2>
        <input
          type="search"
          className="sc-filter"
          placeholder={t('scFilter')}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <div className="sc-list">
          {SHORTCUT_GROUPS.map((group) => {
            const groupRows = rows.filter((row) => row.group === group.id)
            if (groupRows.length === 0) return null
            return (
              <section key={group.id} className="sc-group">
                <h3>{t(group.labelKey)}</h3>
                {groupRows.map((row) => (
                  <div key={row.id} className="sc-row">
                    <span className="sc-label">{row.label}</span>
                    <kbd className="sc-keys">{row.keys}</kbd>
                  </div>
                ))}
              </section>
            )
          })}
          {rows.length === 0 && <p className="sc-empty">{t('scNone')}</p>}
        </div>
        <div className="modal-actions">
          <button className="btn-primary" onClick={onClose}>
            {t('appSettingsCancel')}
          </button>
        </div>
      </div>
    </div>
  )
}

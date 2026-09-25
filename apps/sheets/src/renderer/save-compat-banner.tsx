import { useI18n } from './i18n/locale'
import { saveCompatLabelKey, type SaveCompatFinding } from './save-compat'

interface Props {
  /// Findings currently active in the workbook; rendering is suppressed when
  /// empty (the appear/disappear/dismiss policy lives in App.tsx).
  readonly findings: readonly SaveCompatFinding[]
  readonly onDismiss: () => void
}

/**
 * Pre-flight warning for fail-closed saves (PAR-206): a non-blocking strip
 * between the formula bar and the grid listing the constructs in this
 * workbook that will not survive a save, so the user learns BEFORE investing
 * work instead of at the ⌘S error. Dismissing hides the current finding set
 * for the session; a different set re-shows it. Labels come from the shared
 * registry (save-compat.ts), never from a local copy.
 */
export function SaveCompatBanner({ findings, onDismiss }: Props): React.JSX.Element {
  const { t } = useI18n()
  return (
    // role="status": announced politely, never steals focus — the banner is
    // informational and must not block editing (DoD: non-blocking).
    <section className="save-compat-banner" role="status">
      <span className="save-compat-title">{t('appSaveCompatTitle')}</span>
      <ul className="save-compat-list">
        {findings.map((finding) => (
          <li key={finding.id}>
            {t(saveCompatLabelKey(finding.id))}
            {finding.detail === undefined ? '' : t('appSaveCompatDetail', { item: finding.detail })}
          </li>
        ))}
      </ul>
      <button
        type="button"
        className="save-compat-dismiss"
        data-tip={t('appSaveCompatDismiss')}
        aria-label={t('appSaveCompatDismiss')}
        onClick={onDismiss}
      >
        ✕
      </button>
    </section>
  )
}

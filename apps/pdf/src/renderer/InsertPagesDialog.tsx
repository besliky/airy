/**
 * Insert Pages from PDF dialog: source preview tiles, insert position radio
 * group, and page range. Pure presentation — the position/range state and the
 * confirm validation live in App; the dialog semantics (role, focus trap,
 * Escape from anywhere, focus back to the trigger) come from the shared
 * useModalDialog hook.
 */
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from './i18n/locale'
import type { InsertSourcePageShape } from '../shared/ipc'

type InsertPagesPos = 'front' | 'end' | 'current' | 'after'

/** Insert-source preview tile size: the page shape scaled to fit a 44px box */
const insertTileSize = (shape: InsertSourcePageShape): { width: number; height: number } => {
  const k = 44 / Math.max(shape.width, shape.height, 1)
  return { width: Math.max(10, shape.width * k), height: Math.max(10, shape.height * k) }
}

export function InsertPagesDialog(props: {
  readonly source: { name: string; pages: InsertSourcePageShape[] }
  /** 1-based current page, for the "current position" radio label */
  readonly currentPage: number
  readonly pos: InsertPagesPos
  readonly onPos: (pos: InsertPagesPos) => void
  readonly afterPage: string
  readonly onAfterPage: (value: string) => void
  readonly afterInvalid: boolean
  readonly range: string
  readonly onRange: (value: string) => void
  readonly rangeInvalid: boolean
  /** Insert is running: every control disables and a spinner shows until the
   * rewritten file has reloaded (App keeps the dialog open meanwhile) */
  readonly busy: boolean
  readonly onConfirm: () => void
  readonly onClose: () => void
}) {
  const { t } = useI18n()
  const dialog = useModalDialog(props.onClose)
  const { source, currentPage, pos, onPos, afterPage, onAfterPage, afterInvalid } = props
  const { range, onRange, rangeInvalid, busy, onConfirm, onClose } = props

  return (
    <div className="pdf-modal-mask" {...dialog.backdropProps} onClick={onClose}>
      <div
        className="pdf-modal"
        {...dialog.dialogProps}
        aria-busy={busy}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title" {...dialog.titleProps}>
          {t('insertPdfPagesTitle')}
        </div>
        <div className="pdf-modal-hint">
          {t('insertPdfSourceInfo', { name: source.name, total: source.pages.length })}
        </div>
        <div className="insert-source-pages">
          {source.pages.slice(0, 12).map((shape, i) => {
            const tile = insertTileSize(shape)
            return (
              <div
                key={i}
                className="insert-source-tile"
                style={{ width: tile.width, height: tile.height }}
                title={`${shape.width}×${shape.height}`}
              >
                {i + 1}
              </div>
            )
          })}
          {source.pages.length > 12 && (
            <div
              className="insert-source-tile insert-source-more"
              style={insertTileSize(source.pages[0]!)}
            >
              +{source.pages.length - 12}
            </div>
          )}
        </div>
        <div className="pdf-modal-row">
          <span>{t('insertPdfPosition')}</span>
          <label className="pdf-modal-check">
            <input
              type="radio"
              name="insert-pos"
              checked={pos === 'front'}
              disabled={busy}
              onChange={() => onPos('front')}
            />
            {t('insertPosFront')}
          </label>
          <label className="pdf-modal-check">
            <input
              type="radio"
              name="insert-pos"
              checked={pos === 'current'}
              disabled={busy}
              onChange={() => onPos('current')}
            />
            {t('insertPosCurrent', { page: currentPage })}
          </label>
        </div>
        <div className="pdf-modal-row">
          <span />
          <label className="pdf-modal-check">
            <input
              type="radio"
              name="insert-pos"
              checked={pos === 'after'}
              disabled={busy}
              onChange={() => onPos('after')}
            />
            {t('insertPosAfter')}
            <input
              className={`pdf-modal-input insert-after-page${afterInvalid ? ' invalid' : ''}`}
              value={afterPage}
              disabled={busy || pos !== 'after'}
              onChange={(e) => onAfterPage(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && onConfirm()}
            />
          </label>
          <label className="pdf-modal-check">
            <input
              type="radio"
              name="insert-pos"
              checked={pos === 'end'}
              disabled={busy}
              onChange={() => onPos('end')}
            />
            {t('insertPosEnd')}
          </label>
        </div>
        <input
          className={`pdf-modal-input${rangeInvalid ? ' invalid' : ''}`}
          value={range}
          placeholder={t('insertPdfRangeHint', { total: source.pages.length })}
          autoFocus
          disabled={busy}
          onChange={(e) => onRange(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && onConfirm()}
        />
        {busy && (
          <div className="pdf-modal-busy" role="status">
            {t('insertPdfBusy')}
          </div>
        )}
        <div className="pdf-modal-actions">
          <button className="pdf-modal-btn" disabled={busy} onClick={onClose}>
            {t('cancel')}
          </button>
          <button className="pdf-modal-btn primary" disabled={busy} onClick={onConfirm}>
            {t('ok')}
          </button>
        </div>
      </div>
    </div>
  )
}

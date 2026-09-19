/**
 * Alt text editor (Word's "Edit Alt Text" pane): a title and a description.
 * Shared by the Picture Format, Shape Format and Table Layout ribbons — each
 * writes its own storage (wp:docPr title/descr for pictures and shapes,
 * w:tblCaption/w:tblDescription for tables) through the caller's onApply.
 */
import React, { useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'

export interface AltTextValue {
  title: string
  description: string
}

export function AltTextDialog({
  initial,
  onApply,
  onCancel,
}: {
  initial: AltTextValue
  onApply: (value: AltTextValue) => void
  onCancel: () => void
}) {
  const { t } = useI18n()
  const [value, setValue] = useState<AltTextValue>(initial)
  const titleRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    titleRef.current?.focus()
  }, [])

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onCancel()}>
      <div className="modal alt-text-modal" style={{ width: 360, maxWidth: 'calc(100vw - 32px)' }}>
        <h2>{t('ribbonAltText')}</h2>
        <p className="modal-desc">{t('ribbonAltTextHint')}</p>
        <label className="alt-text-field">
          <span>{t('ribbonAltTitle')}</span>
          <input
            ref={titleRef}
            value={value.title}
            maxLength={255}
            onChange={(e) => setValue((v) => ({ ...v, title: e.target.value }))}
            onKeyDown={(e) => {
              if (e.key === 'Enter') onApply(value)
            }}
          />
        </label>
        <label className="alt-text-field">
          <span>{t('ribbonAltDescription')}</span>
          <textarea
            rows={4}
            value={value.description}
            onChange={(e) => setValue((v) => ({ ...v, description: e.target.value }))}
          />
        </label>
        <div className="modal-actions">
          <button onClick={onCancel}>{t('ribbonCancel')}</button>
          <button className="primary" onClick={() => onApply(value)}>
            {t('ribbonApply')}
          </button>
        </div>
      </div>
    </div>
  )
}

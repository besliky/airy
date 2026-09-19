/**
 * Layout chooser for Export PDF: full pages stay the vector one-slide-per-page
 * export; Notes pages and Handouts (2/3 per sheet) reuse the print sheet's
 * exact page assembly, so the exported file matches what Print produces.
 */
import React, { useState } from 'react'
import { useI18n } from '../i18n/locale'
import type { PdfExportLayout } from '../file-actions'
import { useModalKeys } from './modal-keys'

export function PdfExportDialog({
  slideCount,
  onExport,
  onClose,
}: {
  /** Visible slides in the deck (the export skips hidden ones) */
  slideCount: number
  onExport: (layout: PdfExportLayout) => void
  onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [layout, setLayout] = useState<PdfExportLayout>('full')
  const { ref, onKeyDown } = useModalKeys(onClose)
  const perPage = layout === 'handout2' ? 2 : layout === 'handout3' ? 3 : 1
  const pageCount = Math.ceil(slideCount / perPage)
  const options: Array<[PdfExportLayout, Parameters<typeof t>[0]]> = [
    ['full', 'appPrintLayoutFull'],
    ['notes', 'appPrintLayoutNotes'],
    ['handout2', 'appPrintLayoutHandout2'],
    ['handout3', 'appPrintLayoutHandout3'],
  ]
  return (
    <div className="modal-backdrop" onClick={onClose} onKeyDown={onKeyDown} ref={ref}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonFileExportPdf')}</h2>
        <fieldset>
          <legend>{t('appPrintLayoutGroup')}</legend>
          {options.map(([k, key]) => (
            <label key={k} className="print-radio">
              <input
                type="radio"
                name="pdf-export-layout"
                checked={layout === k}
                onChange={() => setLayout(k)}
              />
              {t(key)}
            </label>
          ))}
        </fieldset>
        <div className="print-page-count">
          {slideCount > 0 ? t('appPrintPageCount', { n: pageCount }) : t('appExportNoSlides')}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('appSettingsCancel')}</button>
          <button className="primary" disabled={slideCount === 0} onClick={() => onExport(layout)}>
            {t('ribbonFileExportPdf')}
          </button>
        </div>
      </div>
    </div>
  )
}

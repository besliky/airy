import type { ReactElement } from 'react'
import type { TFunc } from './i18n/locale'

/** One embedded file of the document (a PDF portfolio child or an attached file),
    as enumerated by pdf.js getAttachments. `id` is the pdf.js name-tree key that
    getAttachmentContent(id) resolves to the bytes. */
export interface PdfAttachment {
  id: string
  filename: string
  description: string
}

/** Read-only attachments list (UX-1734): PDF portfolios carry child documents
    that own no page of the cover file, so without this panel their content is
    invisible — the viewer used to show just the cover page with no indication
    that the document holds anything else. */
export function AttachmentsPanel({
  attachments,
  t,
  onOpen,
}: {
  attachments: PdfAttachment[]
  t: TFunc
  onOpen: (attachment: PdfAttachment) => void
}): ReactElement {
  return (
    <div className="pdf-attachments">
      <div className="pdf-attachments-title">
        {t('attachmentsCount', { n: attachments.length })}
      </div>
      {attachments.map((a) => (
        <div key={a.id} className="pdf-attachment">
          <span className="pdf-attachment-name" data-tip={a.description || a.filename}>
            {a.filename}
          </span>
          <button
            className="pdf-attachment-open"
            data-tip={t('attachmentOpen')}
            aria-label={`${t('attachmentOpen')}: ${a.filename}`}
            onClick={() => onOpen(a)}
          >
            {t('attachmentOpen')}
          </button>
        </div>
      ))}
    </div>
  )
}

import type { TFunc } from './i18n/locale'
import { ocrAvailableOnPlatform } from './ocr-layer'
import { IconScan } from './icons'

/** Quiet status-bar affordance for raster-only documents (UX-1733b). While a
    document has no text layer anywhere, selecting text with the mouse does
    nothing and search cannot hit — the chip names the situation and, per
    platform, says where OCR stands (reuse of the UX-1733 platform split). A
    line in the status bar, not a modal; it disappears once recognition has
    produced text (the caller tracks the effective search index). */
export function ScanNotice({
  show,
  t,
  ocr = ocrAvailableOnPlatform(),
}: {
  show: boolean
  t: TFunc
  /** Platform verdict override (tests); defaults to the live sniff */
  ocr?: boolean
}) {
  if (!show) return null
  return (
    <span
      className="status-item pdf-status-scan"
      data-tip={t(ocr ? 'scanNoTextTipOcr' : 'scanNoTextTipNoOcr')}
    >
      <IconScan size={12} />
      <span className="pdf-status-scan-label">{t('scanNoTextLabel')}</span>
    </span>
  )
}

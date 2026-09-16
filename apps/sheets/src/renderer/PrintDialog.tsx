/**
 * Print dialog modeled on the docs/slides pattern: page preview on the left,
 * page setup options on the right, Print hands the same layout the PDF
 * export builds to the system print dialog. The preview shows the print HTML
 * itself (scaled to the pane — the exact content that will print) with the
 * page count from a headless printToPDF pass, so pagination is exact while
 * the preview pane stays light.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'

import { useI18n } from './i18n/locale'
import type { PrintSetupOverrides } from './page-layout-actions'
import type { WorkbookExportPdfRequest } from '../shared/desktop-api'

/// OOXML paper-size codes the print stack maps to Electron page sizes
/// (print-html's PAPER_SIZES); labels are localized in the app shard.
export const PRINT_PAPER_SIZES: ReadonlyArray<{
  code: number
  labelKey:
    | 'dlgPaperA4'
    | 'dlgPaperA3'
    | 'dlgPaperA5'
    | 'dlgPaperLetter'
    | 'dlgPaperLegal'
    | 'dlgPaperTabloid'
    | 'dlgPaperExecutive'
}> = [
  { code: 9, labelKey: 'dlgPaperA4' },
  { code: 8, labelKey: 'dlgPaperA3' },
  { code: 11, labelKey: 'dlgPaperA5' },
  { code: 1, labelKey: 'dlgPaperLetter' },
  { code: 5, labelKey: 'dlgPaperLegal' },
  { code: 3, labelKey: 'dlgPaperTabloid' },
  { code: 7, labelKey: 'dlgPaperExecutive' },
]

interface EffectiveSetup {
  paperSize: number
  orientation: 'portrait' | 'landscape'
  scale: number
  fitToPage: boolean
}

export function PrintDialog({
  buildRequest,
  onClose,
  setStatus,
}: {
  /// Lays the active sheet out with the dialog's per-job overrides; throws
  /// with a localized message when there is nothing printable.
  readonly buildRequest: (overrides: PrintSetupOverrides) => Promise<{
    request: WorkbookExportPdfRequest
    effective: EffectiveSetup
  }>
  readonly onClose: () => void
  readonly setStatus: ((message: string) => void) | null
}): React.JSX.Element {
  const { t } = useI18n()
  const [previewHtml, setPreviewHtml] = useState<string | null>(null)
  const [pageCount, setPageCount] = useState<number | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [paperSize, setPaperSize] = useState(9)
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait')
  const [scale, setScale] = useState(100)
  const [fitToPage, setFitToPage] = useState(false)
  const [printing, setPrinting] = useState(false)
  const frameRef = useRef<HTMLIFrameElement | null>(null)
  const paneRef = useRef<HTMLDivElement | null>(null)
  const printRequestRef = useRef<WorkbookExportPdfRequest | null>(null)
  const runRef = useRef(0)

  // EffectivePageSetup.scale is a percent (100 = 100%); the payload's
  // computeScale turns it into the request fraction.
  const overrides: PrintSetupOverrides = useMemo(
    () => ({
      paperSize,
      orientation,
      scale: fitToPage ? undefined : scale,
      fitToPage,
    }),
    [paperSize, orientation, scale, fitToPage],
  )

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  // Seed the controls from the sheet's effective saved setup first, then
  // (re)build the preview whenever an option changes.
  const [seeded, setSeeded] = useState(false)
  useEffect(() => {
    if (seeded) return
    let alive = true
    void (async () => {
      try {
        const { effective } = await buildRequest({})
        if (!alive) return
        setPaperSize(effective.paperSize)
        setOrientation(effective.orientation)
        setScale(effective.fitToPage ? 100 : Math.round(effective.scale))
        setFitToPage(effective.fitToPage)
      } catch {
        // Nothing printable: the preview pass below reports the error.
      } finally {
        if (alive) setSeeded(true)
      }
    })()
    return () => {
      alive = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [seeded])

  useEffect(() => {
    if (!seeded) return
    const run = ++runRef.current
    let alive = true
    void (async () => {
      try {
        const built = await buildRequest(overrides)
        if (!alive || run !== runRef.current) return
        printRequestRef.current = built.request
        setPreviewHtml(built.request.html)
        setError(null)
        const preview = await window.desktopApi.previewPrint(built.request)
        if (!alive || run !== runRef.current) return
        if (preview.ok) setPageCount(preview.pageCount)
        else setError(preview.error)
      } catch (reason: unknown) {
        if (alive && run === runRef.current) {
          setPreviewHtml(null)
          setPageCount(null)
          setError(reason instanceof Error ? reason.message : t('appPrintPreviewFailed'))
        }
      }
    })()
    return () => {
      alive = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [seeded, overrides])

  /// Scale the preview document so its page width fits the pane.
  const applyZoom = useCallback(() => {
    const doc = frameRef.current?.contentDocument
    const pane = paneRef.current
    if (!doc || !pane) return
    doc.body.style.zoom = ''
    const pageWidth = doc.body.scrollWidth
    if (pageWidth > 0)
      doc.body.style.zoom = String(Math.min((pane.clientWidth - 44) / pageWidth, 1))
  }, [])

  useEffect(() => {
    const pane = paneRef.current
    if (!pane) return
    const observer = new ResizeObserver(applyZoom)
    observer.observe(pane)
    return () => observer.disconnect()
  }, [applyZoom])

  const doPrint = async (): Promise<void> => {
    const request = printRequestRef.current
    if (!request || printing) return
    setPrinting(true)
    try {
      const result = await window.desktopApi.printWorkbook(request)
      if (result.ok) onClose()
      else if (result.error) setStatus?.(t('appPrintFailed', { error: result.error }))
      // ok=false without an error = canceled in the system dialog: keep open
    } finally {
      setPrinting(false)
    }
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal print-dialog" onClick={(event) => event.stopPropagation()}>
        <h2>{t('dlgPrintTitle')}</h2>
        <div className="print-dialog-body">
          <div className="print-preview-pane" ref={paneRef}>
            {previewHtml !== null ? (
              <iframe
                ref={frameRef}
                title="print-preview"
                sandbox="allow-same-origin"
                srcDoc={previewHtml}
                onLoad={applyZoom}
              />
            ) : (
              <div className="print-preview-empty">{error ?? t('dlgPrintRendering')}</div>
            )}
          </div>
          <div className="print-options">
            <fieldset>
              <legend>{t('dlgPrintPaper')}</legend>
              <select
                className="print-select"
                value={paperSize}
                onChange={(event) => setPaperSize(Number(event.target.value))}
              >
                {PRINT_PAPER_SIZES.map((paper) => (
                  <option key={paper.code} value={paper.code}>
                    {t(paper.labelKey)}
                  </option>
                ))}
              </select>
            </fieldset>
            <fieldset>
              <legend>{t('dlgPrintOrientation')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-orientation"
                  checked={orientation === 'portrait'}
                  onChange={() => setOrientation('portrait')}
                />
                {t('dlgPrintPortrait')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-orientation"
                  checked={orientation === 'landscape'}
                  onChange={() => setOrientation('landscape')}
                />
                {t('dlgPrintLandscape')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('dlgPrintScaling')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-scaling"
                  checked={!fitToPage}
                  onChange={() => setFitToPage(false)}
                />
                {t('dlgPrintScaleAdjust')}
              </label>
              {!fitToPage && (
                <label className="print-radio print-scale-row">
                  <input
                    type="number"
                    className="print-range-input print-scale-input"
                    min={10}
                    max={200}
                    step={5}
                    value={scale}
                    onChange={(event) => {
                      const next = Number(event.target.value)
                      if (Number.isFinite(next)) setScale(Math.min(200, Math.max(10, next)))
                    }}
                  />
                  %
                </label>
              )}
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-scaling"
                  checked={fitToPage}
                  onChange={() => setFitToPage(true)}
                />
                {t('dlgPrintFitSheet')}
              </label>
            </fieldset>
            <div className="print-page-count">
              {pageCount !== null ? t('dlgPrintPageCount', { n: pageCount }) : ''}
            </div>
          </div>
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('dlgCancel')}</button>
          <button
            className="primary"
            disabled={previewHtml === null || printing}
            onClick={() => void doPrint()}
          >
            {printing ? t('dlgPrintProgress') : t('dlgPrintTitle')}
          </button>
        </div>
      </div>
    </div>
  )
}

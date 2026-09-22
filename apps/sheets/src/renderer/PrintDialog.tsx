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
import type { PrintGuard } from './print-guard'
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

/// Preview rebuilds are debounced: rapid option toggles (scrolling the scale
/// slider, flipping orientation) coalesce into one headless printToPDF pass
/// for the last state instead of one full pass per change.
const PREVIEW_DEBOUNCE_MS = 250

interface EffectiveSetup {
  paperSize: number
  orientation: 'portrait' | 'landscape'
  scale: number
  fitToPage: boolean
  pageOrder: 'down-then-over' | 'over-then-down'
}

/// What the job prints (Excel's Print Selection / Active Sheets / Entire
/// Workbook).
export type PrintDialogScope = 'selection' | 'active-sheet' | 'workbook'

/// The Scaling radio: a fixed percent, fit-to-width (all columns on one page
/// across, height unconstrained), or the sheet's saved fit (Excel's "Fit
/// sheet on one page" — the file's fitToWidth/fitToHeight budget).
export type PrintFitMode = 'adjust' | 'fit-width' | 'fit-page'

/// Initial dialog control values for a sheet's effective saved page setup:
/// fit-to-page takes over the scale control at 100%, and the saved page
/// order preselects (pageSetup@pageOrder, Excel-style).
export function controlsFromEffective(effective: EffectiveSetup): {
  paperSize: number
  orientation: 'portrait' | 'landscape'
  scale: number
  fitToPage: boolean
  pageOrder: 'down-then-over' | 'over-then-down'
} {
  return {
    paperSize: effective.paperSize,
    orientation: effective.orientation,
    scale: effective.fitToPage ? 100 : Math.round(effective.scale),
    fitToPage: effective.fitToPage,
    pageOrder: effective.pageOrder,
  }
}

export function PrintDialog({
  buildRequest,
  onClose,
  setStatus,
}: {
  /// Lays the active sheet out with the dialog's per-job overrides; throws
  /// with a localized message when there is nothing printable. Also returns
  /// the print guard: the detected "prints N pages across / foreign paper"
  /// pain of the job as configured (BUG-1603).
  readonly buildRequest: (overrides: PrintSetupOverrides) => Promise<{
    request: WorkbookExportPdfRequest
    effective: EffectiveSetup
    guard: PrintGuard
  }>
  readonly onClose: () => void
  readonly setStatus: ((message: string) => void) | null
}): React.JSX.Element {
  const { t } = useI18n()
  const [previewHtml, setPreviewHtml] = useState<string | null>(null)
  const [pageCount, setPageCount] = useState<number | null>(null)
  const [error, setError] = useState<string | null>(null)
  /// Raw English error from the main process — surfaced only as a tooltip on
  /// the localized message, never as the visible text.
  const [errorDetail, setErrorDetail] = useState<string | null>(null)
  const [paperSize, setPaperSize] = useState(9)
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('portrait')
  const [scale, setScale] = useState(100)
  const [fitMode, setFitMode] = useState<PrintFitMode>('adjust')
  const [scope, setScope] = useState<PrintDialogScope>('active-sheet')
  const [pageOrder, setPageOrder] = useState<'down-then-over' | 'over-then-down'>('down-then-over')
  const [collate, setCollate] = useState(false)
  const [printing, setPrinting] = useState(false)
  /// The pain the guard detected for the job as last built (null until the
  /// first build resolves).
  const [guard, setGuard] = useState<PrintGuard | null>(null)
  /// The user picked a paper themselves: the file's paper no longer speaks
  /// for them, so the paper suggestion stands down.
  const [paperTouched, setPaperTouched] = useState(false)
  const frameRef = useRef<HTMLIFrameElement | null>(null)
  const paneRef = useRef<HTMLDivElement | null>(null)
  const printRequestRef = useRef<WorkbookExportPdfRequest | null>(null)
  const runRef = useRef(0)

  // EffectivePageSetup.scale is a percent (100 = 100%); the payload's
  // computeScale turns it into the request fraction. Fit-to-width applies
  // Excel's "1 page wide by [blank] tall" budget; fit-page keeps whatever
  // the sheet saved (Excel's fit-sheet-on-one-page), exactly as before.
  const overrides: PrintSetupOverrides = useMemo(
    () => ({
      paperSize,
      orientation,
      scale: fitMode === 'adjust' ? scale : undefined,
      fitToPage: fitMode !== 'adjust',
      ...(fitMode === 'fit-width' ? { fitToWidth: 1, fitToHeight: 0 } : {}),
      scope,
      pageOrder,
      collate,
    }),
    [paperSize, orientation, scale, fitMode, scope, pageOrder, collate],
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
        const { effective, guard: seededGuard } = await buildRequest({})
        if (!alive) return
        const seeded = controlsFromEffective(effective)
        setPaperSize(seeded.paperSize)
        setOrientation(seeded.orientation)
        setScale(seeded.scale)
        setFitMode(seeded.fitToPage ? 'fit-page' : 'adjust')
        setPageOrder(seeded.pageOrder)
        setGuard(seededGuard)
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
    let alive = true
    // Debounced: the cleanup of a superseded run cancels its pending timer,
    // so only the last state within the debounce window reaches the main
    // process's printToPDF pass; a run already in flight is cut off by the
    // runRef stamp instead.
    const timer = window.setTimeout(() => {
      const run = ++runRef.current
      void (async () => {
        try {
          const built = await buildRequest(overrides)
          if (!alive || run !== runRef.current) return
          printRequestRef.current = built.request
          setPreviewHtml(built.request.html)
          setGuard(built.guard)
          setError(null)
          setErrorDetail(null)
          const preview = await window.desktopApi.previewPrint(built.request)
          if (!alive || run !== runRef.current) return
          if (preview.ok) setPageCount(preview.pageCount)
          else {
            setError(t('appPrintPreviewFailed'))
            setErrorDetail(preview.error)
          }
        } catch (reason: unknown) {
          if (alive && run === runRef.current) {
            setPreviewHtml(null)
            setPageCount(null)
            setGuard(null)
            setErrorDetail(null)
            // buildRequest throws localized messages when nothing is printable
            setError(reason instanceof Error ? reason.message : t('appPrintPreviewFailed'))
          }
        }
      })()
    }, PREVIEW_DEBOUNCE_MS)
    return () => {
      alive = false
      window.clearTimeout(timer)
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
                title={t('dlgPrintPreviewFrame')}
                aria-label={t('dlgPrintPreviewFrame')}
                sandbox="allow-same-origin"
                srcDoc={previewHtml}
                onLoad={applyZoom}
              />
            ) : (
              <div className="print-preview-empty" title={errorDetail ?? undefined}>
                {error ?? t('dlgPrintRendering')}
              </div>
            )}
          </div>
          <div className="print-options">
            <fieldset>
              <legend>{t('dlgPrintScope')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-scope"
                  checked={scope === 'selection'}
                  onChange={() => setScope('selection')}
                />
                {t('dlgPrintScopeSelection')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-scope"
                  checked={scope === 'active-sheet'}
                  onChange={() => setScope('active-sheet')}
                />
                {t('dlgPrintScopeActiveSheets')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-scope"
                  checked={scope === 'workbook'}
                  onChange={() => setScope('workbook')}
                />
                {t('dlgPrintScopeWorkbook')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('dlgPrintPaper')}</legend>
              <select
                className="print-select"
                value={paperSize}
                onChange={(event) => {
                  setPaperSize(Number(event.target.value))
                  setPaperTouched(true)
                }}
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
                  checked={fitMode === 'adjust'}
                  onChange={() => setFitMode('adjust')}
                />
                {t('dlgPrintScaleAdjust')}
              </label>
              {fitMode === 'adjust' && (
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
                  checked={fitMode === 'fit-width'}
                  onChange={() => setFitMode('fit-width')}
                />
                {t('dlgPrintFitWidth')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-scaling"
                  checked={fitMode === 'fit-page'}
                  onChange={() => setFitMode('fit-page')}
                />
                {t('dlgPrintFitSheet')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('dlgPrintPageOrder')}</legend>
              <select
                className="print-select"
                value={pageOrder}
                onChange={(event) =>
                  setPageOrder(
                    event.target.value === 'over-then-down' ? 'over-then-down' : 'down-then-over',
                  )
                }
              >
                <option value="down-then-over">{t('dlgPrintOrderDownThenOver')}</option>
                <option value="over-then-down">{t('dlgPrintOrderOverThenDown')}</option>
              </select>
              <label className="print-radio">
                <input
                  type="checkbox"
                  checked={collate}
                  onChange={(event) => setCollate(event.target.checked)}
                />
                {t('dlgPrintCollate')}
              </label>
            </fieldset>
            {(guard?.suggestFitToWidth === true ||
              (guard?.suggestPaper === true && !paperTouched)) && (
              <div className="print-guard" role="note">
                {guard?.suggestFitToWidth === true && guard.stripsAcross !== null && (
                  <p>{t('dlgPrintGuardStrips', { n: guard.stripsAcross })}</p>
                )}
                <div className="print-guard-actions">
                  {guard?.suggestFitToWidth === true && (
                    <button type="button" onClick={() => setFitMode('fit-width')}>
                      {t('dlgPrintGuardFitWidth')}
                    </button>
                  )}
                  {guard?.suggestPaper === true && !paperTouched && (
                    <button
                      type="button"
                      onClick={() => {
                        if (guard === null) return
                        setPaperSize(guard.localePaperSize)
                        setPaperTouched(true)
                      }}
                    >
                      {t('dlgPrintGuardPaper', {
                        paper: t(
                          PRINT_PAPER_SIZES.find((p) => p.code === guard?.localePaperSize)
                            ?.labelKey ?? 'dlgPaperA4',
                        ),
                      })}
                    </button>
                  )}
                </div>
              </div>
            )}
            <div className="print-page-count" title={errorDetail ?? undefined}>
              {error ?? (pageCount !== null ? t('dlgPrintPageCount', { n: pageCount }) : '')}
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

/**
 * Options + progress dialog for File > Export Video: resolution (720p/1080p
 * slide height), frame rate, rehearsed-timings pacing with a seconds-per-slide
 * fallback, and transition crossfades. The live duration estimate comes from
 * the same pure timeline the export walks (deck facts fetched once by App).
 *
 * While exporting, the dialog shows the pipeline progress (render slides →
 * record frames) and offers Cancel; on completion (or cancel) it closes —
 * the result surfaces through the status bar like the other exports.
 */
import React, { useMemo, useRef, useState } from 'react'
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from '../i18n/locale'
import type { TransitionSpec } from '../../shared/ipc'
import type { VideoExportPhase, VideoExportSettings } from '../file-actions'
import { buildVideoTimeline, hasRehearseTimings, type VideoPlanSlide } from '../video-plan'
import { pickRecorderMime } from '../video-export'
import { formatClock } from '../slideshow-utils'

export function ExportVideoDialog({
  slides,
  advanceMs,
  transitions,
  onExport,
  onClose,
}: {
  /** Full deck (hidden flags drive the estimate exactly like the export) */
  slides: ReadonlyArray<VideoPlanSlide>
  /** Rehearsed auto-advance ms per deck index (fetched by App once) */
  advanceMs: ReadonlyArray<number | null>
  /** Transition spec per deck index */
  transitions: ReadonlyArray<TransitionSpec>
  onExport: (
    settings: VideoExportSettings,
    onProgress: (phase: VideoExportPhase, done: number, total: number) => void,
    cancel: { current: boolean },
  ) => Promise<boolean>
  onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [heightPreset, setHeightPreset] = useState<720 | 1080>(720)
  const [fps, setFps] = useState(30)
  const deckHasTimings = useMemo(() => hasRehearseTimings(advanceMs), [advanceMs])
  const [useTimings, setUseTimings] = useState(true)
  const [secondsPerSlide, setSecondsPerSlide] = useState(5)
  const [includeTransitions, setIncludeTransitions] = useState(true)
  const [phase, setPhase] = useState<'idle' | VideoExportPhase>('idle')
  const [progress, setProgress] = useState({ done: 0, total: 0 })
  const cancelBox = useRef({ current: false }).current
  // mirrors `phase` for the hook's close callback (Escape must not orphan a run)
  const exportingRef = useRef(false)
  const dialog = useModalDialog(() => {
    if (!exportingRef.current) onClose()
  })

  const mime = useMemo(
    () =>
      typeof MediaRecorder === 'undefined'
        ? null
        : pickRecorderMime((m) => MediaRecorder.isTypeSupported(m)),
    [],
  )

  const effectiveUseTimings = useTimings && deckHasTimings
  const estimate = useMemo(
    () =>
      buildVideoTimeline({
        slides,
        advanceMs,
        transitions,
        options: { fps, useTimings: effectiveUseTimings, secondsPerSlide, includeTransitions },
      }),
    [slides, advanceMs, transitions, fps, effectiveUseTimings, secondsPerSlide, includeTransitions],
  )
  const visibleCount = estimate.items.length

  const start = async () => {
    setPhase('render')
    exportingRef.current = true
    setProgress({ done: 0, total: 0 })
    cancelBox.current = false
    const ok = await onExport(
      { fps, heightPreset, useTimings: effectiveUseTimings, secondsPerSlide, includeTransitions },
      (p, done, total) => {
        setPhase(p)
        setProgress({ done, total })
      },
      cancelBox,
    )
    // close on completion and cancel alike; failures keep the status bar message
    if (ok || cancelBox.current) onClose()
    else {
      exportingRef.current = false
      setPhase('idle')
    }
  }

  const exporting = phase !== 'idle'
  const percent = progress.total > 0 ? Math.round((progress.done / progress.total) * 100) : 0

  return (
    // backdrop click does nothing mid-export: the run is cancelled via the
    // Cancel button, not by orphaning the recording behind a closed dialog
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onClick={exporting ? undefined : onClose}
    >
      <div className="modal" {...dialog.dialogProps} onClick={(e) => e.stopPropagation()}>
        <h2 {...dialog.titleProps}>{t('ribbonFileExportVideo')}</h2>
        {exporting ? (
          <div className="video-export-progress">
            <div className="video-export-progress-label">
              {phase === 'render'
                ? t('appExportVideoRendering', { done: progress.done, total: progress.total })
                : t('appExportVideoRecording', { percent })}
            </div>
            <div
              className="video-export-bar"
              role="progressbar"
              aria-valuemin={0}
              aria-valuemax={100}
              aria-valuenow={percent}
            >
              <div style={{ width: `${percent}%` }} />
            </div>
            <div className="video-export-note">{t('appExportVideoRealtimeNote')}</div>
          </div>
        ) : (
          <>
            <fieldset className="video-export-fieldset">
              <legend>{t('appExportVideoResolution')}</legend>
              <label className="dlg-radio">
                <input
                  type="radio"
                  name="video-export-res"
                  checked={heightPreset === 720}
                  onChange={() => setHeightPreset(720)}
                />
                {t('appExportVideo720p')}
              </label>
              <label className="dlg-radio">
                <input
                  type="radio"
                  name="video-export-res"
                  checked={heightPreset === 1080}
                  onChange={() => setHeightPreset(1080)}
                />
                {t('appExportVideo1080p')}
              </label>
            </fieldset>
            <fieldset className="video-export-fieldset">
              <legend>{t('appExportVideoFps')}</legend>
              <label className="dlg-radio">
                <input
                  type="radio"
                  name="video-export-fps"
                  checked={fps === 30}
                  onChange={() => setFps(30)}
                />
                {t('appExportVideoFps30')}
              </label>
              <label className="dlg-radio">
                <input
                  type="radio"
                  name="video-export-fps"
                  checked={fps === 24}
                  onChange={() => setFps(24)}
                />
                {t('appExportVideoFps24')}
              </label>
            </fieldset>
            <fieldset className="video-export-fieldset">
              <legend>{t('appExportVideoTimingsGroup')}</legend>
              <label className="dlg-check">
                <input
                  type="checkbox"
                  checked={effectiveUseTimings}
                  disabled={!deckHasTimings}
                  onChange={(e) => setUseTimings(e.target.checked)}
                />
                {deckHasTimings ? t('appExportVideoTimings') : t('appExportVideoNoTimings')}
              </label>
              <label className="dlg-check video-export-seconds-row">
                {t('appExportVideoSecondsPerSlide')}
                <input
                  className="video-export-seconds"
                  type="number"
                  min={1}
                  max={60}
                  step={0.5}
                  value={secondsPerSlide}
                  disabled={effectiveUseTimings}
                  onChange={(e) => {
                    const v = Number(e.target.value)
                    if (Number.isFinite(v)) setSecondsPerSlide(Math.min(60, Math.max(1, v)))
                  }}
                />
              </label>
              <label className="dlg-check">
                <input
                  type="checkbox"
                  checked={includeTransitions}
                  onChange={(e) => setIncludeTransitions(e.target.checked)}
                />
                {t('appExportVideoTransitions')}
              </label>
            </fieldset>
            <div className="video-export-summary">
              {mime ? (
                <div>
                  {mime.container === 'mp4' ? t('appExportVideoMp4') : t('appExportVideoWebm')}
                </div>
              ) : (
                <div className="video-export-error">{t('appExportVideoNoEncoder')}</div>
              )}
              <div>
                {visibleCount > 0
                  ? t('appExportVideoDuration', {
                      duration: formatClock(estimate.totalMs),
                      count: visibleCount,
                    })
                  : t('appExportNoSlides')}
              </div>
            </div>
          </>
        )}
        <div className="modal-actions">
          {exporting ? (
            <button onClick={() => (cancelBox.current = true)}>{t('appSettingsCancel')}</button>
          ) : (
            <>
              <button onClick={onClose}>{t('appSettingsCancel')}</button>
              <button
                className="primary"
                disabled={visibleCount === 0 || !mime}
                onClick={() => void start()}
              >
                {t('ribbonFileExportVideo')}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  )
}

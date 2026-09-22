/**
 * PowerPoint "Slide Size → Custom" dialog (Design tab): width/height in cm or
 * inches plus orientation, converted to EMU on OK. Bounds match PowerPoint's
 * slide size limits (0.5in..56in ≈ 1.27cm..142.24cm).
 */
import React, { useState } from 'react'
import { useModalDialog } from '@airy-office/ui'
import { useI18n } from '../i18n/locale'

/** EMU per unit (1 in = 2.54 cm = 914400 EMU; 1 cm = 360000 EMU). */
const EMU_PER_CM = 360000
const CM_PER_IN = 2.54
/** PowerPoint's slide size bounds: 0.5in (1.27cm) .. 56in (142.24cm). */
export const MIN_SLIDE_CM = 1.27
export const MAX_SLIDE_CM = 142.24

export type LengthUnit = 'cm' | 'in'

/** Convert a length in `unit` to EMU. */
export function toEmu(value: number, unit: LengthUnit): number {
  return Math.round(value * (unit === 'cm' ? EMU_PER_CM : EMU_PER_CM * CM_PER_IN))
}

/** Length bound in the given unit, for input validation. */
function bounds(unit: LengthUnit): { min: number; max: number } {
  const k = unit === 'cm' ? 1 : 1 / CM_PER_IN
  return { min: MIN_SLIDE_CM * k, max: MAX_SLIDE_CM * k }
}

/** Clamp a length (cm) into PowerPoint's bounds. */
export function clampLengthCm(cm: number): number {
  return Math.min(MAX_SLIDE_CM, Math.max(MIN_SLIDE_CM, cm))
}

export function SlideSizeDialog({
  widthCm,
  heightCm,
  onApply,
  onClose,
}: {
  /** Current slide size in cm (initial field values and orientation) */
  widthCm: number
  heightCm: number
  onApply: (cxEmu: number, cyEmu: number) => void
  onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const dialog = useModalDialog(onClose)
  const landscape = widthCm >= heightCm
  const [unit, setUnit] = useState<LengthUnit>('cm')
  const [long, setLong] = useState(String(round2(landscape ? widthCm : heightCm)))
  const [short, setShort] = useState(String(round2(landscape ? heightCm : widthCm)))
  const [isLandscape, setIsLandscape] = useState(landscape)

  const lv = parseFloat(long)
  const sv = parseFloat(short)
  const b = bounds(unit)
  const inRange = (v: number) => Number.isFinite(v) && v >= b.min && v <= b.max
  const valid = inRange(lv) && inRange(sv)
  const numeric = Number.isFinite(lv) && Number.isFinite(sv)

  // Unit switch converts the entered values in place so nothing is lost
  const switchUnit = (next: LengthUnit) => {
    if (next === unit) return
    const k = next === 'in' ? 1 / CM_PER_IN : CM_PER_IN
    const conv = (v: string) => String(Math.round((parseFloat(v) || 0) * k * 100) / 100)
    setLong(conv(long))
    setShort(conv(short))
    setUnit(next)
  }

  const apply = () => {
    if (!valid) return
    const longCm = clampLengthCm(unit === 'cm' ? lv : lv * CM_PER_IN)
    const shortCm = clampLengthCm(unit === 'cm' ? sv : sv * CM_PER_IN)
    onApply(
      toEmu(isLandscape ? longCm : shortCm, 'cm'),
      toEmu(isLandscape ? shortCm : longCm, 'cm'),
    )
    onClose()
  }

  return (
    <div className="modal-backdrop" {...dialog.backdropProps} onClick={onClose}>
      <div
        className="modal slide-size-dlg"
        {...dialog.dialogProps}
        onClick={(e) => e.stopPropagation()}
      >
        <h2 {...dialog.titleProps}>{t('ribbonSlideSize')}</h2>
        <div className="slide-size-fields">
          <label>
            <span>{isLandscape ? t('paneSizeWidth') : t('paneSizeHeight')}</span>
            <input
              type="number"
              min={0}
              step={0.1}
              value={long}
              onChange={(e) => setLong(e.target.value)}
            />
          </label>
          <label>
            <span>{isLandscape ? t('paneSizeHeight') : t('paneSizeWidth')}</span>
            <input
              type="number"
              min={0}
              step={0.1}
              value={short}
              onChange={(e) => setShort(e.target.value)}
            />
          </label>
          <div className="slide-size-unit" role="radiogroup" aria-label={t('ribbonSlideSize')}>
            {(['cm', 'in'] as const).map((u) => (
              <button
                key={u}
                className={unit === u ? 'on' : ''}
                onClick={() => switchUnit(u)}
                aria-pressed={unit === u}
              >
                {u}
              </button>
            ))}
          </div>
        </div>
        <fieldset>
          <legend>{t('appPrintOrientation')}</legend>
          <label className="print-radio">
            <input
              type="radio"
              name="slide-size-orientation"
              checked={isLandscape}
              onChange={() => setIsLandscape(true)}
            />
            {t('appPrintLandscape')}
          </label>
          <label className="print-radio">
            <input
              type="radio"
              name="slide-size-orientation"
              checked={!isLandscape}
              onChange={() => setIsLandscape(false)}
            />
            {t('appPrintPortrait')}
          </label>
        </fieldset>
        {numeric && !valid && <div className="slide-size-error">{t('appSlideSizeRange')}</div>}
        <div className="modal-actions">
          <button onClick={onClose}>{t('appSettingsCancel')}</button>
          <button className="primary" disabled={!valid} onClick={apply}>
            {t('paneOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Two-decimal round (cm fields match PowerPoint's 0.01cm precision). */
function round2(v: number): number {
  return Math.round(v * 100) / 100
}

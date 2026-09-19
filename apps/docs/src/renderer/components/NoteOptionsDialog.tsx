import { useState } from 'react'
import type { NoteNumberFmt, NoteNumbering } from '@airy-office/docx-engine'
import { Dropdown, useModalDialog } from '@airy-office/ui'
import { noteNumberingOfKind } from '../note-format'
import { useI18n, type StringKey } from '../i18n/locale'

const FORMAT_KEYS: Array<{ key: NoteNumberFmt; nameKey: StringKey }> = [
  { key: 'decimal', nameKey: 'refsFmtDecimal' },
  { key: 'lowerLetter', nameKey: 'refsFmtLowerLetter' },
  { key: 'upperLetter', nameKey: 'refsFmtUpperLetter' },
  { key: 'lowerRoman', nameKey: 'refsFmtLowerRoman' },
  { key: 'upperRoman', nameKey: 'refsFmtUpperRoman' },
]

const CUSTOM = 'custom' as const

/**
 * Word's Footnote and Endnote dialog: per-kind numbering format (1/a/i or a
 * fixed custom mark), start-at, whole-document footnote↔endnote conversion and
 * single-note conversion of the reference under the caret.
 */
export function NoteOptionsDialog({
  value,
  hasSelection,
  selectionKind,
  onApply,
  onConvert,
  onClose,
}: {
  value: { footnotes?: NoteNumbering; endnotes?: NoteNumbering }
  /** a docNoteRef sits at the caret (enables "convert this note") */
  hasSelection: boolean
  /** kind of that reference, when any */
  selectionKind?: 'footnote' | 'endnote'
  onApply: (next: { footnotes?: NoteNumbering; endnotes?: NoteNumbering }) => void
  onConvert: (from: 'footnote' | 'endnote', which: 'all' | 'current') => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [kind, setKind] = useState<'footnote' | 'endnote'>('footnote')
  const current = noteNumberingOfKind(kind, value)
  const [fmt, setFmt] = useState<string>(
    current?.numFmt ?? (kind === 'footnote' ? 'decimal' : 'lowerRoman'),
  )
  const [customMark, setCustomMark] = useState(current?.customMark ?? '')
  const [startAt, setStartAt] = useState(String(current?.numStart ?? 1))
  const dialog = useModalDialog(onClose)

  /**
   * BUG-1005: switching the Footnotes/Endnotes radio reloads the chosen
   * kind's numbering into the controls (Word's dialog swaps the shown
   * settings with the radio). Without the re-seed, OK silently overwrote
   * the other kind's options with values seeded from the kind that was
   * selected when the dialog opened.
   */
  const pickKind = (next: 'footnote' | 'endnote') => {
    setKind(next)
    const seeding = noteNumberingOfKind(next, value)
    setFmt(seeding?.numFmt ?? (next === 'footnote' ? 'decimal' : 'lowerRoman'))
    setCustomMark(seeding?.customMark ?? '')
    setStartAt(String(seeding?.numStart ?? 1))
  }

  const submit = () => {
    const model: NoteNumbering =
      fmt === CUSTOM
        ? { numFmt: 'decimal', customMark: customMark.trim().slice(0, 2) || '*' }
        : { numFmt: fmt as NoteNumberFmt }
    const start = Math.min(9999, Math.max(1, Math.round(Number(startAt) || 1)))
    if (start !== 1) model.numStart = start
    const next: { footnotes?: NoteNumbering; endnotes?: NoteNumbering } = {
      ...value,
      ...(kind === 'footnote' ? { footnotes: model } : { endnotes: model }),
    }
    onApply(next)
    onClose()
  }

  const convert = (from: 'footnote' | 'endnote', which: 'all' | 'current') => {
    onConvert(from, which)
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('refsNoteOptionsTitle')}</h2>
        <div className="modal-row ln-modes">
          <label className="ln-radio">
            <input
              type="radio"
              name="refs-kind"
              checked={kind === 'footnote'}
              onChange={() => pickKind('footnote')}
            />
            {t('refsFootnotes')}
          </label>
          <label className="ln-radio">
            <input
              type="radio"
              name="refs-kind"
              checked={kind === 'endnote'}
              onChange={() => pickKind('endnote')}
            />
            {t('refsEndnotes')}
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('refsNumberFormat')}
            <Dropdown
              value={fmt}
              ariaLabel={t('refsNumberFormat')}
              options={[
                ...FORMAT_KEYS.map((f) => ({ value: f.key, label: t(f.nameKey) })),
                { value: CUSTOM, label: t('refsFmtCustomMark') },
              ]}
              onPick={setFmt}
            />
          </label>
          {fmt === CUSTOM ? (
            <label>
              {t('refsCustomMark')}
              <input
                value={customMark}
                onChange={(e) => setCustomMark(e.target.value)}
                placeholder="*"
                maxLength={2}
                onKeyDown={(e) => e.key === 'Enter' && submit()}
              />
            </label>
          ) : (
            <label>
              {t('refsStartAt')}
              <input
                type="number"
                min={1}
                max={9999}
                step={1}
                value={startAt}
                onChange={(e) => setStartAt(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && submit()}
              />
            </label>
          )}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={() => convert('footnote', 'all')}>
            {t('refsConvertToEndnotes')}
          </button>
          <button className="btn-ghost" onClick={() => convert('endnote', 'all')}>
            {t('refsConvertToFootnotes')}
          </button>
          <button
            className="btn-ghost"
            disabled={!hasSelection || selectionKind === undefined}
            title={
              hasSelection && selectionKind
                ? t(
                    selectionKind === 'footnote'
                      ? 'refsConvertThisEndnote'
                      : 'refsConvertThisFootnote',
                  )
                : t('refsConvertNoneHint')
            }
            onClick={() => selectionKind && convert(selectionKind, 'current')}
          >
            {t('refsConvertSelected')}
          </button>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

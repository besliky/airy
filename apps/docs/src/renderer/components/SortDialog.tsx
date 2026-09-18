/**
 * Word's Sort dialog (Home ▸ Sort): reorder the selected table's rows by a
 * column key, or the selected paragraphs as lines. One "Then by" level rides
 * the primary key; tables with vertically merged cells are refused up front
 * (the sort would orphan the merge continuations).
 */
import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import { Dropdown, type DropdownOption, useModalDialog } from '@airy-office/ui'
import { useI18n } from '../i18n/locale'
import {
  sortScope,
  sortSelectedParagraphs,
  sortTableRows,
  type SortFieldType,
  type SortLevel,
} from '../editor/sort'

interface LevelState {
  column: number
  type: SortFieldType
  descending: boolean
}

const NONE = 'none'

export function SortDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const dialog = useModalDialog(onClose)
  // computed once per open: the dialog describes the selection it was opened on
  const [scope] = useState(() => sortScope(editor.state))
  const table = scope?.kind === 'table' ? scope : null
  const paragraphs = scope?.kind === 'paragraphs' ? scope : null

  const [primary, setPrimary] = useState<LevelState>({ column: 0, type: 'text', descending: false })
  const [thenBy, setThenBy] = useState<LevelState | null>(null)
  const [headerRow, setHeaderRow] = useState(table?.headerRow ?? false)

  /** Word names columns by their header text once a header row exists */
  const columnLabel = (column: number): string =>
    (headerRow && table?.headerLabels[column]) || t('appSortColumn', { n: column + 1 })

  const columnOptions: DropdownOption[] = Array.from(
    { length: table?.columnCount ?? 0 },
    (_, column) => ({ value: String(column), label: columnLabel(column) }),
  )

  const typeOptions: Array<{ value: SortFieldType; label: string }> = [
    { value: 'text', label: t('appSortTypeText') },
    { value: 'number', label: t('appSortTypeNumber') },
    { value: 'date', label: t('appSortTypeDate') },
  ]
  const directionOptions = [
    { value: 'asc', label: t('appSortAscending') },
    { value: 'desc', label: t('appSortDescending') },
  ]

  const sortableRows = (table?.rowCount ?? 0) - (headerRow ? 1 : 0)
  const refused = table?.hasVerticalMerge === true
  const canApply =
    !!scope &&
    !refused &&
    ((table !== null && sortableRows >= 2 && columnOptions.length > 0) ||
      (paragraphs !== null && paragraphs.count >= 2))

  const apply = () => {
    if (!editor.isEditable) {
      onClose()
      return
    }
    if (!canApply) return
    const levels: SortLevel[] = [primary, ...(thenBy !== null && table ? [thenBy] : [])]
    if (table) {
      sortTableRows({ levels, headerRow })(editor.state, editor.view.dispatch)
    } else if (paragraphs) {
      sortSelectedParagraphs(levels)(editor.state, editor.view.dispatch)
    }
    onClose()
  }

  /** one Sort-by / Then-by row: column (tables only) + type + direction */
  const levelRow = (
    label: string,
    state: LevelState | null,
    pick: (next: LevelState | null) => void,
    withNone: boolean,
  ) => {
    const current = state ?? { column: 0, type: 'text' as SortFieldType, descending: false }
    return (
      <div className="font-dialog-row">
        {table !== null && (
          <label>
            {label}
            <Dropdown
              value={state === null && withNone ? NONE : String(current.column)}
              ariaLabel={label}
              options={
                withNone
                  ? [{ value: NONE, label: t('appSortNone') }, ...columnOptions]
                  : columnOptions
              }
              onPick={(v) => pick(v === NONE ? null : { ...current, column: Number(v) })}
            />
          </label>
        )}
        <label>
          {t('appSortType')}
          <Dropdown
            value={current.type}
            ariaLabel={t('appSortType')}
            options={typeOptions}
            onPick={(v) => pick({ ...current, type: v })}
            disabled={state === null}
          />
        </label>
        <label>
          {t('appSortAscending')}/{t('appSortDescending')}
          <Dropdown
            value={current.descending ? 'desc' : 'asc'}
            ariaLabel={t('appSortAscending')}
            options={directionOptions}
            onPick={(v) => pick({ ...current, descending: v === 'desc' })}
            disabled={state === null}
          />
        </label>
      </div>
    )
  }

  return (
    <div
      className="modal-backdrop"
      {...dialog.backdropProps}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal" {...dialog.dialogProps}>
        <h2 {...dialog.titleProps}>{t('appSortDialogTitle')}</h2>
        <p className="sort-scope">
          {table
            ? t('appSortScopeTable', { rows: table.rowCount, cols: table.columnCount })
            : paragraphs
              ? t('appSortScopeParagraphs', { n: paragraphs.count })
              : t('appSortNoScope')}
        </p>
        {refused && <p className="sort-warn">{t('appSortMergedUnsupported')}</p>}
        {levelRow(t('appSortBy'), primary, (next) => next && setPrimary(next), false)}
        {table !== null && levelRow(t('appSortThenBy'), thenBy, setThenBy, true)}
        {table !== null && (
          <label className="font-check">
            <input
              type="checkbox"
              checked={headerRow}
              onChange={(e) => setHeaderRow(e.target.checked)}
            />
            {t('appSortHeaderRow')}
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply} disabled={!canApply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

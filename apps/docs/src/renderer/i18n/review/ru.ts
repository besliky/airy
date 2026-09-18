import type { zh } from './zh'

export const ru = {
  reviewCompareMerge: 'Сравнить (объединить как правки)',
  reviewCompareMergeDesc: 'Показать различия как исправления, которые можно принять или отклонить',
  reviewComparePanel: 'Только панель различий',
  reviewComparePanelDesc: 'Показать различия по абзацам в боковой панели, не изменяя документ',
  reviewCompareMerged:
    'Сравнение с {name}: {added} вставок, {removed} удалений и {changed} изменений объединено как рецензирование',
  reviewCompareMergedApprox:
    'Сравнение с {name}: {added} вставок, {removed} удалений и {changed} изменений объединено как рецензирование (документы слишком велики для точного выравнивания абзацев)',
  reviewCompareIdentical: 'Различий с {name} нет: документы идентичны',
  reviewCompareDegraded:
    'Документы слишком велики для точного выравнивания абзацев: различия сопоставлены по позиции',
  reviewComparePendingRevisions:
    'В документе есть непринятые исправления. Сначала примите или отклоните их, затем сравнивайте снова',
  reviewCompareReadonly:
    'Для сравнения (объединения как правки) документ должен быть редактируемым; текущий документ доступен только для чтения',
  reviewComparing: 'Сравнение…',
} satisfies Record<keyof typeof zh, string>

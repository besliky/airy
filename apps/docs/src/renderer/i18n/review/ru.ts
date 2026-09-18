import type { zh } from './zh'

export const ru = {
  reviewCompareMerge: 'Сравнить (объединить как правки)',
  reviewComparePanel: 'Только панель различий',
  reviewCompareMerged:
    'Сравнение с {name}: {added} вставок, {removed} удалений и {changed} изменений объединено как рецензирование',
  reviewCompareIdentical: 'Различий с {name} нет: документы идентичны',
} satisfies Record<keyof typeof zh, string>

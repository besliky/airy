import type { zh } from './zh'

export const pl = {
  reviewCompareMerge: 'Porównaj (scal jako śledzone zmiany)',
  reviewComparePanel: 'Pokaż tylko panel różnic',
  reviewCompareMerged:
    'Porównano z {name}: scalono {added} wstawień, {removed} usunięć i {changed} zmian jako śledzone zmiany',
  reviewCompareIdentical: 'Brak różnic z {name}: dokumenty są identyczne',
} satisfies Record<keyof typeof zh, string>

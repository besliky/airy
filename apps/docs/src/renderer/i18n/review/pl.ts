import type { zh } from './zh'

export const pl = {
  reviewCompareMerge: 'Porównaj (scal jako śledzone zmiany)',
  reviewComparePanel: 'Pokaż tylko panel różnic',
  reviewCompareMerged:
    'Porównano z {name}: scalono {added} wstawień, {removed} usunięć i {changed} zmian jako śledzone zmiany',
  reviewCompareIdentical: 'Brak różnic z {name}: dokumenty są identyczne',
  reviewCompareMergedApprox:
    'Porównano z {name}: scalono {added} wstawień, {removed} usunięć i {changed} zmian jako śledzone zmiany (dokumenty zbyt duże do dokładnego dopasowania akapitów)',
  reviewCompareDegraded:
    'Dokumenty są zbyt duże do dokładnego dopasowania akapitów: różnice sparowano według pozycji',
} satisfies Record<keyof typeof zh, string>

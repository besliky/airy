import type { zh } from './zh'

export const pl = {
  reviewCompareMerge: 'Porównaj (scal jako śledzone zmiany)',
  reviewCompareMergeDesc:
    'Pokazuje różnice jako śledzone zmiany, które można zaakceptować lub odrzucić',
  reviewComparePanel: 'Pokaż tylko panel różnic',
  reviewComparePanelDesc:
    'Wyświetla różnice między akapitami w panelu bocznym bez modyfikowania dokumentu',
  reviewCompareMerged:
    'Porównano z {name}: scalono {added} wstawień, {removed} usunięć i {changed} zmian jako śledzone zmiany',
  reviewCompareIdentical: 'Brak różnic z {name}: dokumenty są identyczne',
  reviewCompareMergedApprox:
    'Porównano z {name}: scalono {added} wstawień, {removed} usunięć i {changed} zmian jako śledzone zmiany (dokumenty zbyt duże do dokładnego dopasowania akapitów)',
  reviewCompareDegraded:
    'Dokumenty są zbyt duże do dokładnego dopasowania akapitów: różnice sparowano według pozycji',
  reviewComparePendingRevisions:
    'Dokument zawiera oczekujące śledzone zmiany. Zaakceptuj lub odrzuć je przed ponownym porównaniem',
  reviewCompareReadonly:
    'Porównanie (scalenie jako śledzone zmiany) wymaga edytowalnego dokumentu; ten dokument jest tylko do odczytu',
} satisfies Record<keyof typeof zh, string>

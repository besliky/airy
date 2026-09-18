import type { zh } from './zh'

export const cs = {
  reviewCompareMerge: 'Porovnat (sloučit jako sledované změny)',
  reviewCompareMergeDesc:
    'Zobrazí rozdíly jako sledované změny, které můžete přijmout nebo odmítnout',
  reviewComparePanel: 'Zobrazit pouze panel rozdílů',
  reviewComparePanelDesc: 'Vypíše rozdíly odstavců v postranním podokně bez úprav dokumentu',
  reviewCompareMerged:
    'Porovnáno s {name}: {added} vložení, {removed} odstranění a {changed} změny sloučeny jako sledované změny',
  reviewCompareIdentical: 'Žádné rozdíly s {name}: dokumenty jsou identické',
  reviewCompareMergedApprox:
    'Porovnáno s {name}: {added} vložení, {removed} odstranění a {changed} změny sloučeny jako sledované změny (dokumenty příliš velké pro přesné párování odstavců)',
  reviewCompareDegraded:
    'Dokumenty jsou příliš velké pro přesné párování odstavců: rozdíly byly spárovány podle pozice',
  reviewComparePendingRevisions:
    'Dokument obsahuje nevyřízené sledované změny. Před novým porovnáním je přijměte nebo odmítněte',
  reviewCompareReadonly:
    'Porovnání (sloučit jako sledované změny) vyžaduje upravitelný dokument; tento dokument je jen pro čtení',
  reviewComparing: 'Porovnávání…',
} satisfies Record<keyof typeof zh, string>

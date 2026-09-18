import type { zh } from './zh'

export const cs = {
  reviewCompareMerge: 'Porovnat (sloučit jako sledované změny)',
  reviewComparePanel: 'Zobrazit pouze panel rozdílů',
  reviewCompareMerged:
    'Porovnáno s {name}: {added} vložení, {removed} odstranění a {changed} změny sloučeny jako sledované změny',
  reviewCompareIdentical: 'Žádné rozdíly s {name}: dokumenty jsou identické',
  reviewCompareMergedApprox:
    'Porovnáno s {name}: {added} vložení, {removed} odstranění a {changed} změny sloučeny jako sledované změny (dokumenty příliš velké pro přesné párování odstavců)',
  reviewCompareDegraded:
    'Dokumenty jsou příliš velké pro přesné párování odstavců: rozdíly byly spárovány podle pozice',
} satisfies Record<keyof typeof zh, string>

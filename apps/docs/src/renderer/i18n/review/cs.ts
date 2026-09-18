import type { zh } from './zh'

export const cs = {
  reviewCompareMerge: 'Porovnat (sloučit jako sledované změny)',
  reviewComparePanel: 'Zobrazit pouze panel rozdílů',
  reviewCompareMerged:
    'Porovnáno s {name}: {added} vložení, {removed} odstranění a {changed} změny sloučeny jako sledované změny',
  reviewCompareIdentical: 'Žádné rozdíly s {name}: dokumenty jsou identické',
} satisfies Record<keyof typeof zh, string>

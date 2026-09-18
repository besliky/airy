import type { zh } from './zh'

export const fr = {
  reviewCompareMerge: 'Comparer (fusionner en révisions)',
  reviewComparePanel: 'Afficher seulement le volet des différences',
  reviewCompareMerged:
    'Comparé avec {name} : {added} insertions, {removed} suppressions et {changed} modifications fusionnées en révisions',
  reviewCompareIdentical: 'Aucune différence avec {name} : les documents sont identiques',
} satisfies Record<keyof typeof zh, string>

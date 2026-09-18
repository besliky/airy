import type { zh } from './zh'

export const fr = {
  reviewCompareMerge: 'Comparer (fusionner en révisions)',
  reviewComparePanel: 'Afficher seulement le volet des différences',
  reviewCompareMerged:
    'Comparé avec {name} : {added} insertions, {removed} suppressions et {changed} modifications fusionnées en révisions',
  reviewCompareIdentical: 'Aucune différence avec {name} : les documents sont identiques',
  reviewCompareMergedApprox:
    'Comparé avec {name} : {added} insertions, {removed} suppressions et {changed} modifications fusionnées en révisions (documents trop volumineux pour un appariement exact des paragraphes)',
  reviewCompareDegraded:
    'Les documents sont trop volumineux pour un appariement exact des paragraphes : les différences ont été appariées par position',
  reviewComparePendingRevisions:
    'Le document contient des révisions en attente. Acceptez-les ou rejetez-les avant de comparer à nouveau',
} satisfies Record<keyof typeof zh, string>

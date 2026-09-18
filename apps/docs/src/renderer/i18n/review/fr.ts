import type { zh } from './zh'

export const fr = {
  reviewCompareMerge: 'Comparer (fusionner en révisions)',
  reviewCompareMergeDesc:
    'Affiche les différences sous forme de modifications que vous pouvez accepter ou refuser',
  reviewComparePanel: 'Afficher seulement le volet des différences',
  reviewComparePanelDesc:
    'Liste les différences de paragraphes dans un volet latéral sans modifier le document',
  reviewCompareMerged:
    'Comparé avec {name} : {added} insertions, {removed} suppressions et {changed} modifications fusionnées en révisions',
  reviewCompareIdentical: 'Aucune différence avec {name} : les documents sont identiques',
  reviewCompareMergedApprox:
    'Comparé avec {name} : {added} insertions, {removed} suppressions et {changed} modifications fusionnées en révisions (documents trop volumineux pour un appariement exact des paragraphes)',
  reviewCompareDegraded:
    'Les documents sont trop volumineux pour un appariement exact des paragraphes : les différences ont été appariées par position',
  reviewComparePendingRevisions:
    'Le document contient des révisions en attente. Acceptez-les ou rejetez-les avant de comparer à nouveau',
  reviewCompareReadonly:
    'Comparer (fusionner en révisions) nécessite un document modifiable ; ce document est en lecture seule',
} satisfies Record<keyof typeof zh, string>

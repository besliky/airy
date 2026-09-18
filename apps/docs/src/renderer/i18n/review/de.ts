import type { zh } from './zh'

export const de = {
  reviewCompareMerge: 'Vergleichen (als Änderungen zusammenführen)',
  reviewCompareMergeDesc:
    'Zeigt die Unterschiede als nachverfolgte Änderungen, die Sie annehmen oder ablehnen können',
  reviewComparePanel: 'Nur Unterschiedsbereich anzeigen',
  reviewComparePanelDesc:
    'Listet die Absatzunterschiede in einem Seitenbereich, ohne das Dokument zu ändern',
  reviewCompareMerged:
    'Mit {name} verglichen: {added} Einfügungen, {removed} Löschungen und {changed} Änderungen als nachverfolgte Änderungen zusammengeführt',
  reviewCompareIdentical: 'Keine Unterschiede zu {name}: Die Dokumente sind identisch',
  reviewCompareMergedApprox:
    'Mit {name} verglichen: {added} Einfügungen, {removed} Löschungen und {changed} Änderungen als nachverfolgte Änderungen zusammengeführt (Dokumente zu groß für exakten Absatzabgleich)',
  reviewCompareDegraded:
    'Die Dokumente sind für einen exakten Absatzabgleich zu groß: Unterschiede wurden positionell gepaart',
  reviewComparePendingRevisions:
    'Das Dokument enthält offene Änderungen. Nehmen Sie sie an oder lehnen Sie sie ab, bevor Sie erneut vergleichen',
  reviewCompareReadonly:
    'Vergleichen (als Änderungen zusammenführen) erfordert ein bearbeitbares Dokument; dieses Dokument ist schreibgeschützt',
} satisfies Record<keyof typeof zh, string>

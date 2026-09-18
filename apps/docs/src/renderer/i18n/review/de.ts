import type { zh } from './zh'

export const de = {
  reviewCompareMerge: 'Vergleichen (als Änderungen zusammenführen)',
  reviewComparePanel: 'Nur Unterschiedsbereich anzeigen',
  reviewCompareMerged:
    'Mit {name} verglichen: {added} Einfügungen, {removed} Löschungen und {changed} Änderungen als nachverfolgte Änderungen zusammengeführt',
  reviewCompareIdentical: 'Keine Unterschiede zu {name}: Die Dokumente sind identisch',
  reviewCompareMergedApprox:
    'Mit {name} verglichen: {added} Einfügungen, {removed} Löschungen und {changed} Änderungen als nachverfolgte Änderungen zusammengeführt (Dokumente zu groß für exakten Absatzabgleich)',
  reviewCompareDegraded:
    'Die Dokumente sind für einen exakten Absatzabgleich zu groß: Unterschiede wurden positionell gepaart',
} satisfies Record<keyof typeof zh, string>

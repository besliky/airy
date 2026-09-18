import type { zh } from './zh'

export const it = {
  reviewCompareMerge: 'Confronta (unisci come revisioni)',
  reviewComparePanel: 'Mostra solo il riquadro delle differenze',
  reviewCompareMerged:
    'Confrontato con {name}: {added} inserimenti, {removed} eliminazioni e {changed} modifiche uniti come revisioni',
  reviewCompareIdentical: 'Nessuna differenza con {name}: i documenti sono identici',
  reviewCompareMergedApprox:
    'Confrontato con {name}: {added} inserimenti, {removed} eliminazioni e {changed} modifiche uniti come revisioni (documenti troppo grandi per un allineamento esatto dei paragrafi)',
  reviewCompareDegraded:
    'I documenti sono troppo grandi per un allineamento esatto dei paragrafi: le differenze sono state abbinate per posizione',
} satisfies Record<keyof typeof zh, string>

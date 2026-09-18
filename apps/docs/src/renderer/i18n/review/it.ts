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
  reviewComparePendingRevisions:
    'Il documento contiene revisioni in sospeso. Accettale o rifiutale prima di confrontare di nuovo',
  reviewCompareReadonly:
    'Confronta (unisci come revisioni) richiede un documento modificabile; questo documento è in sola lettura',
} satisfies Record<keyof typeof zh, string>

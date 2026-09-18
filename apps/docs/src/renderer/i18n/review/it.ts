import type { zh } from './zh'

export const it = {
  reviewCompareMerge: 'Confronta (unisci come revisioni)',
  reviewCompareMergeDesc: 'Mostra le differenze come modifiche registrate da accettare o rifiutare',
  reviewComparePanel: 'Mostra solo il riquadro delle differenze',
  reviewComparePanelDesc:
    'Elencare le differenze tra paragrafi in un riquadro laterale senza modificare il documento',
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
  reviewComparing: 'Confronto in corso…',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const it = {
  reviewCompareMerge: 'Confronta (unisci come revisioni)',
  reviewComparePanel: 'Mostra solo il riquadro delle differenze',
  reviewCompareMerged:
    'Confrontato con {name}: {added} inserimenti, {removed} eliminazioni e {changed} modifiche uniti come revisioni',
  reviewCompareIdentical: 'Nessuna differenza con {name}: i documenti sono identici',
} satisfies Record<keyof typeof zh, string>

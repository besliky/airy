import type { zh } from './zh'

export const es = {
  reviewCompareMerge: 'Comparar (fusionar como cambios)',
  reviewComparePanel: 'Mostrar solo el panel de diferencias',
  reviewCompareMerged:
    'Comparado con {name}: {added} inserciones, {removed} eliminaciones y {changed} cambios fusionados como cambios controlados',
  reviewCompareIdentical: 'Sin diferencias con {name}: los documentos son idénticos',
  reviewCompareMergedApprox:
    'Comparado con {name}: {added} inserciones, {removed} eliminaciones y {changed} cambios fusionados como cambios controlados (documentos demasiado grandes para una coincidencia exacta de párrafos)',
  reviewCompareDegraded:
    'Los documentos son demasiado grandes para una coincidencia exacta de párrafos: las diferencias se emparejaron por posición',
  reviewComparePendingRevisions:
    'El documento tiene cambios controlados pendientes. Acéptelos o rechácelos antes de volver a comparar',
  reviewCompareReadonly:
    'Comparar (fusionar como cambios) necesita un documento editable; este documento es de solo lectura',
} satisfies Record<keyof typeof zh, string>

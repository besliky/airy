import type { zh } from './zh'

export const es = {
  reviewCompareMerge: 'Comparar (fusionar como cambios)',
  reviewComparePanel: 'Mostrar solo el panel de diferencias',
  reviewCompareMerged:
    'Comparado con {name}: {added} inserciones, {removed} eliminaciones y {changed} cambios fusionados como cambios controlados',
  reviewCompareIdentical: 'Sin diferencias con {name}: los documentos son idénticos',
} satisfies Record<keyof typeof zh, string>

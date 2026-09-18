import type { zh } from './zh'

export const pt = {
  reviewCompareMerge: 'Comparar (mesclar como alterações)',
  reviewComparePanel: 'Mostrar apenas o painel de diferenças',
  reviewCompareMerged:
    'Comparado com {name}: {added} inserções, {removed} exclusões e {changed} alterações mescladas como alterações controladas',
  reviewCompareIdentical: 'Sem diferenças com {name}: os documentos são idênticos',
  reviewCompareMergedApprox:
    'Comparado com {name}: {added} inserções, {removed} exclusões e {changed} alterações mescladas como alterações controladas (documentos grandes demais para correspondência exata de parágrafos)',
  reviewCompareDegraded:
    'Os documentos são grandes demais para correspondência exata de parágrafos: as diferenças foram pareadas por posição',
} satisfies Record<keyof typeof zh, string>

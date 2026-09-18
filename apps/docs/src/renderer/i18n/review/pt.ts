import type { zh } from './zh'

export const pt = {
  reviewCompareMerge: 'Comparar (mesclar como alterações)',
  reviewCompareMergeDesc: 'Mostra as diferenças como alterações controladas a aceitar ou rejeitar',
  reviewComparePanel: 'Mostrar apenas o painel de diferenças',
  reviewComparePanelDesc:
    'Listar as diferenças de parágrafos num painel lateral sem alterar o documento',
  reviewCompareMerged:
    'Comparado com {name}: {added} inserções, {removed} exclusões e {changed} alterações mescladas como alterações controladas',
  reviewCompareIdentical: 'Sem diferenças com {name}: os documentos são idênticos',
  reviewCompareMergedApprox:
    'Comparado com {name}: {added} inserções, {removed} exclusões e {changed} alterações mescladas como alterações controladas (documentos grandes demais para correspondência exata de parágrafos)',
  reviewCompareDegraded:
    'Os documentos são grandes demais para correspondência exata de parágrafos: as diferenças foram pareadas por posição',
  reviewComparePendingRevisions:
    'O documento tem alterações controladas pendentes. Aceite-as ou rejeite-as antes de comparar novamente',
  reviewCompareReadonly:
    'Comparar (mesclar como alterações) requer um documento editável; este documento é somente leitura',
  reviewComparing: 'A comparar…',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const nl = {
  reviewCompareMerge: 'Vergelijken (samenvoegen als wijzigingen)',
  reviewCompareMergeDesc:
    'Toont de verschillen als bijgehouden wijzigingen die u kunt accepteren of weigeren',
  reviewComparePanel: 'Alleen het verschillenpaneel weergeven',
  reviewComparePanelDesc:
    'Somt de alineaverschillen op in een zijpaneel zonder het document te wijzigen',
  reviewCompareMerged:
    'Vergeleken met {name}: {added} invoegingen, {removed} verwijderingen en {changed} wijzigingen samengevoegd als bijgehouden wijzigingen',
  reviewCompareIdentical: 'Geen verschillen met {name}: de documenten zijn identiek',
  reviewCompareMergedApprox:
    'Vergeleken met {name}: {added} invoegingen, {removed} verwijderingen en {changed} wijzigingen samengevoegd als bijgehouden wijzigingen (documenten te groot voor exacte alinea-afstemming)',
  reviewCompareDegraded:
    'De documenten zijn te groot voor exacte alinea-afstemming: verschillen zijn per positie gekoppeld',
  reviewComparePendingRevisions:
    'Het document bevat openstaande wijzigingen. Accepteer of wijs ze af voordat u opnieuw vergelijkt',
  reviewCompareReadonly:
    'Vergelijken (samenvoegen als wijzigingen) vereist een bewerkbaar document; dit document is alleen-lezen',
  reviewComparing: 'Vergelijken…',
} satisfies Record<keyof typeof zh, string>

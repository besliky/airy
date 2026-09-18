import type { zh } from './zh'

export const nl = {
  reviewCompareMerge: 'Vergelijken (samenvoegen als wijzigingen)',
  reviewComparePanel: 'Alleen het verschillenpaneel weergeven',
  reviewCompareMerged:
    'Vergeleken met {name}: {added} invoegingen, {removed} verwijderingen en {changed} wijzigingen samengevoegd als bijgehouden wijzigingen',
  reviewCompareIdentical: 'Geen verschillen met {name}: de documenten zijn identiek',
  reviewCompareMergedApprox:
    'Vergeleken met {name}: {added} invoegingen, {removed} verwijderingen en {changed} wijzigingen samengevoegd als bijgehouden wijzigingen (documenten te groot voor exacte alinea-afstemming)',
  reviewCompareDegraded:
    'De documenten zijn te groot voor exacte alinea-afstemming: verschillen zijn per positie gekoppeld',
  reviewComparePendingRevisions:
    'Het document bevat openstaande wijzigingen. Accepteer of wijs ze af voordat u opnieuw vergelijkt',
} satisfies Record<keyof typeof zh, string>

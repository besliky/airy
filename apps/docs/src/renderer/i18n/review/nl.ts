import type { zh } from './zh'

export const nl = {
  reviewCompareMerge: 'Vergelijken (samenvoegen als wijzigingen)',
  reviewComparePanel: 'Alleen het verschillenpaneel weergeven',
  reviewCompareMerged:
    'Vergeleken met {name}: {added} invoegingen, {removed} verwijderingen en {changed} wijzigingen samengevoegd als bijgehouden wijzigingen',
  reviewCompareIdentical: 'Geen verschillen met {name}: de documenten zijn identiek',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const ms = {
  reviewCompareMerge: 'Bandingkan (gabung sebagai perubahan dijejak)',
  reviewComparePanel: 'Tunjukkan panel perbezaan sahaja',
  reviewCompareMerged:
    'Dibandingkan dengan {name}: {added} sisipan, {removed} pemadaman dan {changed} perubahan digabungkan sebagai perubahan dijejak',
  reviewCompareIdentical: 'Tiada perbezaan dengan {name}: dokumen adalah sama',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const ms = {
  reviewCompareMerge: 'Bandingkan (gabung sebagai perubahan dijejak)',
  reviewCompareMergeDesc:
    'Memaparkan perbezaan sebagai perubahan dijejak yang boleh diterima atau ditolak',
  reviewComparePanel: 'Tunjukkan panel perbezaan sahaja',
  reviewComparePanelDesc:
    'Menyenaraikan perbezaan perenggan dalam panel sisi tanpa mengubah dokumen',
  reviewCompareMerged:
    'Dibandingkan dengan {name}: {added} sisipan, {removed} pemadaman dan {changed} perubahan digabungkan sebagai perubahan dijejak',
  reviewCompareIdentical: 'Tiada perbezaan dengan {name}: dokumen adalah sama',
  reviewCompareMergedApprox:
    'Dibandingkan dengan {name}: {added} sisipan, {removed} pemadaman dan {changed} perubahan digabungkan sebagai perubahan dijejak (dokumen terlalu besar untuk padanan perenggan tepat)',
  reviewCompareDegraded:
    'Dokumen terlalu besar untuk padanan perenggan tepat: perbezaan dipasangkan mengikut kedudukan',
  reviewComparePendingRevisions:
    'Dokumen mengandungi perubahan dijejak yang tertunda. Terima atau tolak dahulu sebelum membanding semula',
  reviewCompareReadonly:
    'Bandingkan (gabung sebagai perubahan dijejak) memerlukan dokumen yang boleh diedit; dokumen ini baca sahaja',
  reviewComparing: 'Membanding…',
} satisfies Record<keyof typeof zh, string>

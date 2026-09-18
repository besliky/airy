import type { zh } from './zh'

export const id = {
  reviewCompareMerge: 'Bandingkan (gabungkan sebagai perubahan terlacak)',
  reviewComparePanel: 'Hanya tampilkan panel perbedaan',
  reviewCompareMerged:
    'Dibandingkan dengan {name}: {added} penyisipan, {removed} penghapusan, dan {changed} perubahan digabungkan sebagai perubahan terlacak',
  reviewCompareIdentical: 'Tidak ada perbedaan dengan {name}: dokumen identik',
} satisfies Record<keyof typeof zh, string>

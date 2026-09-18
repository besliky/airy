import type { zh } from './zh'

export const id = {
  reviewCompareMerge: 'Bandingkan (gabungkan sebagai perubahan terlacak)',
  reviewComparePanel: 'Hanya tampilkan panel perbedaan',
  reviewCompareMerged:
    'Dibandingkan dengan {name}: {added} penyisipan, {removed} penghapusan, dan {changed} perubahan digabungkan sebagai perubahan terlacak',
  reviewCompareIdentical: 'Tidak ada perbedaan dengan {name}: dokumen identik',
  reviewCompareMergedApprox:
    'Dibandingkan dengan {name}: {added} penyisipan, {removed} penghapusan, dan {changed} perubahan digabungkan sebagai perubahan terlacak (dokumen terlalu besar untuk pencocokan paragraf yang tepat)',
  reviewCompareDegraded:
    'Dokumen terlalu besar untuk pencocokan paragraf yang tepat: perbedaan dipasangkan berdasarkan posisi',
  reviewComparePendingRevisions:
    'Dokumen memiliki perubahan terlacak yang tertunda. Terima atau tolak sebelum membandingkan lagi',
} satisfies Record<keyof typeof zh, string>

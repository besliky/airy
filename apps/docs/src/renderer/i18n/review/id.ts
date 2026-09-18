import type { zh } from './zh'

export const id = {
  reviewCompareMerge: 'Bandingkan (gabungkan sebagai perubahan terlacak)',
  reviewCompareMergeDesc:
    'Menampilkan perbedaan sebagai perubahan terlacak yang bisa diterima atau ditolak',
  reviewComparePanel: 'Hanya tampilkan panel perbedaan',
  reviewComparePanelDesc: 'Mendaftarkan perbedaan paragraf di panel samping tanpa mengubah dokumen',
  reviewCompareMerged:
    'Dibandingkan dengan {name}: {added} penyisipan, {removed} penghapusan, dan {changed} perubahan digabungkan sebagai perubahan terlacak',
  reviewCompareIdentical: 'Tidak ada perbedaan dengan {name}: dokumen identik',
  reviewCompareMergedApprox:
    'Dibandingkan dengan {name}: {added} penyisipan, {removed} penghapusan, dan {changed} perubahan digabungkan sebagai perubahan terlacak (dokumen terlalu besar untuk pencocokan paragraf yang tepat)',
  reviewCompareDegraded:
    'Dokumen terlalu besar untuk pencocokan paragraf yang tepat: perbedaan dipasangkan berdasarkan posisi',
  reviewComparePendingRevisions:
    'Dokumen memiliki perubahan terlacak yang tertunda. Terima atau tolak sebelum membandingkan lagi',
  reviewCompareReadonly:
    'Bandingkan (gabungkan sebagai perubahan terlacak) membutuhkan dokumen yang dapat diedit; dokumen ini hanya-baca',
} satisfies Record<keyof typeof zh, string>

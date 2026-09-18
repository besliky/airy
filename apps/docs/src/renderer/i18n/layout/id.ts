import type { zh } from './zh'

export const id = {
  layoutHyphenation: 'Pemenggalan kata',
  layoutHyphNone: 'Tidak ada',
  layoutHyphManual: 'Manual',
  layoutHyphAutomatic: 'Otomatis',
  layoutHyphManualDesc: 'Sisipkan tanda hubung opsional di tempat pemenggalan diinginkan',
  layoutHyphManualHint: 'Pemenggalan manual: tekan {keys} untuk menyisipkan tanda hubung opsional',
  layoutHyphSet: 'Pemenggalan otomatis diaktifkan',
  layoutHyphUnset: 'Pemenggalan otomatis dimatikan',
  layoutSoftHyphen: 'Tanda hubung opsional',
  layoutColsDialogTitle: 'Kolom',
  layoutColsMore: 'Kolom lainnya…',
  layoutColOne: 'Satu',
  layoutColTwo: 'Dua',
  layoutColThree: 'Tiga',
  layoutColLeft: 'Kiri',
  layoutColRight: 'Kanan',
  layoutColSpacing: 'Jarak',
  layoutColWidth: 'Lebar',
  layoutColWidth1: 'Lebar kolom 1',
  layoutLineBetween: 'Garis pemisah kolom',
} satisfies Record<keyof typeof zh, string>

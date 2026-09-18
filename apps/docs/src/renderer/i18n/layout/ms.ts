import type { zh } from './zh'

export const ms = {
  layoutHyphenation: 'Pemisahan suku kata',
  layoutHyphNone: 'Tiada',
  layoutHyphManual: 'Manual',
  layoutHyphAutomatic: 'Automatik',
  layoutHyphManualDesc: 'Sisipkan tanda sempang pilihan di tempat pemisahan dikehendaki',
  layoutHyphManualHint: 'Pemisahan manual: tekan {keys} untuk menyisipkan tanda sempang pilihan',
  layoutHyphSet: 'Pemisahan automatik diaktifkan',
  layoutHyphUnset: 'Pemisahan automatik dimatikan',
  layoutSoftHyphen: 'Tanda sempang pilihan',
  layoutColsDialogTitle: 'Lajur',
  layoutColsMore: 'Lajur lanjutan…',
  layoutColOne: 'Satu',
  layoutColTwo: 'Dua',
  layoutColThree: 'Tiga',
  layoutColLeft: 'Kiri',
  layoutColRight: 'Kanan',
  layoutColSpacing: 'Jarak',
  layoutColWidth: 'Lebar',
  layoutColWidth1: 'Lebar lajur 1',
  layoutLineBetween: 'Garis pemisah lajur',
} satisfies Record<keyof typeof zh, string>

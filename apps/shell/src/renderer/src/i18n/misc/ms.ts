import type { zh } from './zh'

export const ms = {
  today: 'Hari ini',
  yesterday: 'Semalam',
  closeTab: 'Tutup tab',
  errFileNotFound: 'Fail tidak wujud atau telah dipindahkan.',
  errPermissionDenied: 'Akses kepada fail ini ditolak.',
  errIsADirectory: 'Ini ialah folder, bukan fail dokumen.',
  errFileLocked: 'Fail dibuka dalam program lain. Tutup dan cuba lagi.',
  errTooManyFiles: 'Terlalu banyak fail dibuka. Cuba lagi sebentar lagi.',
  tabList: 'Semua tab',
  newTab: 'Tab baharu',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const id = {
  today: 'Hari ini',
  yesterday: 'Kemarin',
  daysAgo: '{n} hari lalu',
  closeTab: 'Tutup tab',
  errFileNotFound: 'Berkas tidak ada atau telah dipindahkan.',
  errPermissionDenied: 'Akses ke berkas ini ditolak.',
  errFileLocked: 'Berkas sedang dibuka di program lain. Tutup lalu coba lagi.',
  errTooManyFiles: 'Terlalu banyak berkas terbuka. Coba lagi sebentar lagi.',
  tabList: 'Semua tab',
  newTab: 'Tab baru',
} satisfies Record<keyof typeof zh, string>

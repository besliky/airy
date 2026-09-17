import type { zh } from './zh'

export const pl = {
  today: 'Dzisiaj',
  yesterday: 'Wczoraj',
  daysAgo: '{n} dni temu',
  closeTab: 'Zamknij kartę',
  errFileNotFound: 'Plik nie istnieje lub został przeniesiony.',
  errPermissionDenied: 'Brak dostępu do tego pliku.',
  errFileLocked: 'Plik jest otwarty w innym programie. Zamknij go i spróbuj ponownie.',
  errTooManyFiles: 'Otwarto zbyt wiele plików. Spróbuj ponownie za chwilę.',
  tabList: 'Wszystkie karty',
  newTab: 'Nowa karta',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const de = {
  today: 'Heute',
  yesterday: 'Gestern',
  closeTab: 'Tab schließen',
  errFileNotFound: 'Die Datei ist nicht vorhanden oder wurde verschoben.',
  errPermissionDenied: 'Zugriff auf diese Datei verweigert.',
  errFileLocked:
    'Die Datei ist in einem anderen Programm geöffnet. Schließen Sie es und versuchen Sie es erneut.',
  errTooManyFiles: 'Zu viele Dateien sind geöffnet. Versuchen Sie es bald erneut.',
  tabList: 'Alle Tabs',
  newTab: 'Neuer Tab',
} satisfies Record<keyof typeof zh, string>

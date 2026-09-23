import type { zh } from './zh'

export const it = {
  today: 'Oggi',
  yesterday: 'Ieri',
  closeTab: 'Chiudi scheda',
  errFileNotFound: 'Il file non esiste o è stato spostato.',
  errPermissionDenied: 'Accesso negato a questo file.',
  errIsADirectory: 'Questa è una cartella, non un file di documento.',
  errFileLocked: 'Il file è aperto in un altro programma. Chiudilo e riprova.',
  errTooManyFiles: 'Troppi file aperti. Riprova tra poco.',
  tabList: 'Tutte le schede',
  newTab: 'Nuova scheda',
} satisfies Record<keyof typeof zh, string>

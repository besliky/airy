import type { zh } from './zh'

export const nl = {
  today: 'Vandaag',
  yesterday: 'Gisteren',
  closeTab: 'Tabblad sluiten',
  errFileNotFound: 'Het bestand bestaat niet of is verplaatst.',
  errPermissionDenied: 'Toegang tot dit bestand geweigerd.',
  errFileLocked: 'Het bestand is geopend in een ander programma. Sluit het en probeer opnieuw.',
  errTooManyFiles: 'Er zijn te veel bestanden open. Probeer het zo opnieuw.',
  tabList: 'Alle tabbladen',
  newTab: 'Nieuw tabblad',
} satisfies Record<keyof typeof zh, string>

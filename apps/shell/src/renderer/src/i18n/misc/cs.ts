import type { zh } from './zh'

export const cs = {
  today: 'Dnes',
  yesterday: 'Včera',
  closeTab: 'Zavřít kartu',
  errFileNotFound: 'Soubor neexistuje nebo byl přesunut.',
  errPermissionDenied: 'Přístup k tomuto souboru byl odepřen.',
  errIsADirectory: 'Toto je složka, nikoli soubor dokumentu.',
  errFileLocked: 'Soubor je otevřený v jiném programu. Zavřete ho a zkuste to znovu.',
  errTooManyFiles: 'Je otevřeno příliš mnoho souborů. Zkuste to za chvíli znovu.',
  tabList: 'Všechny karty',
  newTab: 'Nová karta',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const it = {
  layoutHyphenation: 'Sillabazione',
  layoutHyphNone: 'Nessuna',
  layoutHyphManual: 'Manuale',
  layoutHyphAutomatic: 'Automatica',
  layoutHyphManualDesc: 'Inserisci trattini facoltativi dove vuoi gli a capo',
  layoutHyphManualHint: 'Sillabazione manuale: premi {keys} per inserire un trattino facoltativo',
  layoutHyphSet: 'Sillabazione automatica attivata',
  layoutHyphUnset: 'Sillabazione automatica disattivata',
  layoutSoftHyphen: 'Trattino facoltativo',
  layoutColsDialogTitle: 'Colonne',
  layoutColsMore: 'Altre colonne…',
  layoutColOne: 'Una',
  layoutColTwo: 'Due',
  layoutColThree: 'Tre',
  layoutColLeft: 'Sinistra',
  layoutColRight: 'Destra',
  layoutColSpacing: 'Spaziatura',
  layoutColWidth: 'Larghezza',
  layoutColWidth1: 'Larghezza colonna 1',
  layoutLineBetween: 'Linea tra le colonne',
} satisfies Record<keyof typeof zh, string>

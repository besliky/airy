import type { zh } from './zh'

export const de = {
  layoutHyphenation: 'Silbentrennung',
  layoutHyphNone: 'Keine',
  layoutHyphManual: 'Manuell',
  layoutHyphAutomatic: 'Automatisch',
  layoutHyphManualDesc: 'Fügen Sie bedingte Trennstriche ein, wo Umbrüche gewünscht sind',
  layoutHyphManualHint: 'Manuelle Silbentrennung: {keys} fügt einen bedingten Trennstrich ein',
  layoutHyphSet: 'Automatische Silbentrennung aktiviert',
  layoutHyphUnset: 'Automatische Silbentrennung deaktiviert',
  layoutSoftHyphen: 'Bedingter Trennstrich',
  layoutColsDialogTitle: 'Spalten',
  layoutColsMore: 'Weitere Spalten…',
  layoutColOne: 'Eins',
  layoutColTwo: 'Zwei',
  layoutColThree: 'Drei',
  layoutColLeft: 'Links',
  layoutColRight: 'Rechts',
  layoutColSpacing: 'Abstand',
  layoutColWidth: 'Breite',
  layoutColWidth1: 'Breite Spalte 1',
  layoutLineBetween: 'Trennlinie',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const pl = {
  layoutHyphenation: 'Dzielenie wyrazów',
  layoutHyphNone: 'Brak',
  layoutHyphManual: 'Ręczne',
  layoutHyphAutomatic: 'Automatycznie',
  layoutHyphManualDesc: 'Wstaw miękkie łączniki tam, gdzie mają być złamania',
  layoutHyphManualHint: 'Ręczne dzielenie: naciśnij {keys}, aby wstawić miękki łącznik',
  layoutHyphSet: 'Automatyczne dzielenie wyrazów włączone',
  layoutHyphUnset: 'Automatyczne dzielenie wyrazów wyłączone',
  layoutSoftHyphen: 'Miękki łącznik',
  layoutColsDialogTitle: 'Kolumny',
  layoutColsMore: 'Więcej kolumn…',
  layoutColOne: 'Jedna',
  layoutColTwo: 'Dwie',
  layoutColThree: 'Trzy',
  layoutColLeft: 'Lewa',
  layoutColRight: 'Prawa',
  layoutColSpacing: 'Odstęp (cm)',
  layoutColWidth: 'Szerokość (cm)',
  layoutColWidth1: 'Szerokość kolumny 1 (cm)',
  layoutLineBetween: 'Linia między kolumnami',
} satisfies Record<keyof typeof zh, string>

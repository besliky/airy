import type { zh } from './zh'

export const cs = {
  layoutHyphenation: 'Dělení slov',
  layoutHyphNone: 'Žádné',
  layoutHyphManual: 'Ruční',
  layoutHyphAutomatic: 'Automatické',
  layoutHyphManualDesc: 'Vkládejte měkké spojovníky tam, kde mají být zlomy',
  layoutHyphManualHint: 'Ruční dělení: stiskněte {keys} a vložte měkký spojovník',
  layoutHyphSet: 'Automatické dělení slov zapnuto',
  layoutHyphUnset: 'Automatické dělení slov vypnuto',
  layoutSoftHyphen: 'Měkký spojovník',
  layoutColsDialogTitle: 'Sloupce',
  layoutColsMore: 'Další sloupce…',
  layoutColOne: 'Jeden',
  layoutColTwo: 'Dva',
  layoutColThree: 'Tři',
  layoutColLeft: 'Vlevo',
  layoutColRight: 'Vpravo',
  layoutColSpacing: 'Mezera (cm)',
  layoutColWidth: 'Šířka (cm)',
  layoutColWidth1: 'Šířka sloupce 1 (cm)',
  layoutLineBetween: 'Linka mezi sloupci',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const nl = {
  layoutHyphenation: 'Afbreking',
  layoutHyphNone: 'Geen',
  layoutHyphManual: 'Handmatig',
  layoutHyphAutomatic: 'Automatisch',
  layoutHyphManualDesc: 'Voeg optionele afbreektekens in waar regeleinden moeten komen',
  layoutHyphManualHint:
    'Handmatige afbreking: druk op {keys} om een optioneel afbreekteken in te voegen',
  layoutHyphSet: 'Automatische afbreking ingeschakeld',
  layoutHyphUnset: 'Automatische afbreking uitgeschakeld',
  layoutSoftHyphen: 'Optioneel afbreekteken',
  layoutColsDialogTitle: 'Kolommen',
  layoutColsMore: 'Meer kolommen…',
  layoutColOne: 'Eén',
  layoutColTwo: 'Twee',
  layoutColThree: 'Drie',
  layoutColLeft: 'Links',
  layoutColRight: 'Rechts',
  layoutColSpacing: 'Tussenruimte (cm)',
  layoutColWidth: 'Breedte (cm)',
  layoutColWidth1: 'Breedte kolom 1 (cm)',
  layoutLineBetween: 'Lijn tussen kolommen',
} satisfies Record<keyof typeof zh, string>

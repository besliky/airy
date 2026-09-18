import type { zh } from './zh'

export const he = {
  layoutHyphenation: 'מיקוף',
  layoutHyphNone: 'ללא',
  layoutHyphManual: 'ידני',
  layoutHyphAutomatic: 'אוטומטי',
  layoutHyphManualDesc: 'הוסף מיקפים רכים במקומות שבהם רוצים שבירת שורה',
  layoutHyphManualHint: 'מיקוף ידני: הקש {keys} להוספת מיקף רך',
  layoutHyphSet: 'מיקוף אוטומטי הופעל',
  layoutHyphUnset: 'מיקוף אוטומטי הופסק',
  layoutSoftHyphen: 'מיקף רך',
  layoutColsDialogTitle: 'עמודות',
  layoutColsMore: 'עמודות נוספות…',
  layoutColOne: 'אחת',
  layoutColTwo: 'שתיים',
  layoutColThree: 'שלוש',
  layoutColLeft: 'שמאל',
  layoutColRight: 'ימין',
  layoutColSpacing: 'ריווח',
  layoutColWidth: 'רוחב',
  layoutColWidth1: 'רוחב עמודה 1',
  layoutLineBetween: 'קו בין עמודות',
} satisfies Record<keyof typeof zh, string>

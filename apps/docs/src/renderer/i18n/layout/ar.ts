import type { zh } from './zh'

export const ar = {
  layoutHyphenation: 'وصل الكلمات',
  layoutHyphNone: 'بدون',
  layoutHyphManual: 'يدوي',
  layoutHyphAutomatic: 'تلقائي',
  layoutHyphManualDesc: 'أدرج واصلات اختيارية حيث تريد فواصل الأسطر',
  layoutHyphManualHint: 'وصل يدوي: اضغط {keys} لإدراج واصلة اختيارية',
  layoutHyphSet: 'تم تشغيل الوصل التلقائي للكلمات',
  layoutHyphUnset: 'تم إيقاف الوصل التلقائي للكلمات',
  layoutSoftHyphen: 'واصلة اختيارية',
  layoutColsDialogTitle: 'الأعمدة',
  layoutColsMore: 'أعمدة إضافية…',
  layoutColOne: 'واحد',
  layoutColTwo: 'اثنان',
  layoutColThree: 'ثلاثة',
  layoutColLeft: 'يسار',
  layoutColRight: 'يمين',
  layoutColSpacing: 'التباعد',
  layoutColWidth: 'العرض',
  layoutColWidth1: 'عرض العمود 1',
  layoutLineBetween: 'خط بين الأعمدة',
} satisfies Record<keyof typeof zh, string>

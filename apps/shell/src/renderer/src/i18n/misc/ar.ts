import type { zh } from './zh'

export const ar = {
  today: 'اليوم',
  yesterday: 'أمس',
  daysAgo: 'قبل {n} أيام',
  closeTab: 'إغلاق علامة التبويب',
  errFileNotFound: 'الملف غير موجود أو تم نقله.',
  errPermissionDenied: 'تم رفض الوصول إلى هذا الملف.',
  errFileLocked: 'الملف مفتوح في برنامج آخر. أغلقه ثم أعد المحاولة.',
  errTooManyFiles: 'عدد الملفات المفتوحة كبير جدًا. أعد المحاولة بعد قليل.',
  tabList: 'كل علامات التبويب',
  newTab: 'علامة تبويب جديدة',
} satisfies Record<keyof typeof zh, string>

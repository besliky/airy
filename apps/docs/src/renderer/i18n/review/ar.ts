import type { zh } from './zh'

export const ar = {
  reviewCompareMerge: 'مقارنة (دمج كتغييرات متتبعة)',
  reviewComparePanel: 'إظهار لوحة الفروقات فقط',
  reviewCompareMerged:
    'تمت المقارنة مع {name}: تم دمج {added} إدراجًا و{removed} حذفًا و{changed} تغييرًا كتغييرات متتبعة',
  reviewCompareIdentical: 'لا توجد فروقات مع {name}: المستندان متطابقان',
  reviewCompareMergedApprox:
    'تمت المقارنة مع {name}: تم دمج {added} إدراجًا و{removed} حذفًا و{changed} تغييرًا كتغييرات متتبعة (المستندات كبيرة جدًا لمطابقة الفقرات بدقة)',
  reviewCompareDegraded:
    'المستندات كبيرة جدًا لمطابقة الفقرات بدقة: تم إقران الاختلافات حسب الموضع',
  reviewComparePendingRevisions:
    'يحتوي المستند على تغييرات متتبعة معلقة. اقبلها أو ارفضها قبل المقارنة مرة أخرى',
} satisfies Record<keyof typeof zh, string>

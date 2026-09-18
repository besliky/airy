import type { zh } from './zh'

export const ar = {
  reviewCompareMerge: 'مقارنة (دمج كتغييرات متتبعة)',
  reviewComparePanel: 'إظهار لوحة الفروقات فقط',
  reviewCompareMerged:
    'تمت المقارنة مع {name}: تم دمج {added} إدراجًا و{removed} حذفًا و{changed} تغييرًا كتغييرات متتبعة',
  reviewCompareIdentical: 'لا توجد فروقات مع {name}: المستندان متطابقان',
} satisfies Record<keyof typeof zh, string>

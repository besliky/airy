import type { zh } from './zh'

export const he = {
  reviewCompareMerge: 'השווה (מזג כשינויים מסומנים)',
  reviewComparePanel: 'הצג את חלונית ההבדלים בלבד',
  reviewCompareMerged:
    'הושווה עם {name}: {added} הוספות, {removed} מחיקות ו{changed} שינויים מוזגו כשינויים מסומנים',
  reviewCompareIdentical: 'אין הבדלים עם {name}: המסמכים זהים',
  reviewCompareMergedApprox:
    'הושווה עם {name}: {added} הוספות, {removed} מחיקות ו{changed} שינויים מוזגו כשינויים מסומנים (המסמכים גדולים מדי להתאמת פסקאות מדויקת)',
  reviewCompareDegraded: 'המסמכים גדולים מדי להתאמת פסקאות מדויקת: ההבדלים נוצוו לפי מיקום',
  reviewComparePendingRevisions:
    'במסמך יש שינויים מסומנים ממתינים. קבל או דחה אותם לפני השוואה נוספת',
} satisfies Record<keyof typeof zh, string>

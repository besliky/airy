import type { zh } from './zh'

export const he = {
  reviewCompareMerge: 'השווה (מזג כשינויים מסומנים)',
  reviewComparePanel: 'הצג את חלונית ההבדלים בלבד',
  reviewCompareMerged:
    'הושווה עם {name}: {added} הוספות, {removed} מחיקות ו{changed} שינויים מוזגו כשינויים מסומנים',
  reviewCompareIdentical: 'אין הבדלים עם {name}: המסמכים זהים',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const he = {
  reviewCompareMerge: 'השווה (מזג כשינויים מסומנים)',
  reviewCompareMergeDesc: 'מציג את ההבדלים כשינויים מסומנים שניתן לקבל או לדחות',
  reviewComparePanel: 'הצג את חלונית ההבדלים בלבד',
  reviewComparePanelDesc: 'מציג את הבדלי הפסקאות בחלונית צדדית מבלי לשנות את המסמך',
  reviewCompareMerged:
    'הושווה עם {name}: {added} הוספות, {removed} מחיקות ו{changed} שינויים מוזגו כשינויים מסומנים',
  reviewCompareIdentical: 'אין הבדלים עם {name}: המסמכים זהים',
  reviewCompareMergedApprox:
    'הושווה עם {name}: {added} הוספות, {removed} מחיקות ו{changed} שינויים מוזגו כשינויים מסומנים (המסמכים גדולים מדי להתאמת פסקאות מדויקת)',
  reviewCompareDegraded: 'המסמכים גדולים מדי להתאמת פסקאות מדויקת: ההבדלים נוצוו לפי מיקום',
  reviewComparePendingRevisions:
    'במסמך יש שינויים מסומנים ממתינים. קבל או דחה אותם לפני השוואה נוספת',
  reviewCompareReadonly:
    'השוואה (מזג כשינויים מסומנים) דורשת מסמך הניתן לעריכה; מסמך זה הוא לקריאה בלבד',
} satisfies Record<keyof typeof zh, string>

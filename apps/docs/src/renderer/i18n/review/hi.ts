import type { zh } from './zh'

export const hi = {
  reviewCompareMerge: 'तुलना करें (ट्रैक किए गए परिवर्तनों के रूप में मर्ज करें)',
  reviewComparePanel: 'केवल अंतर पैनल दिखाएँ',
  reviewCompareMerged:
    '{name} के साथ तुलना की गई: {added} जोड़, {removed} हटाने और {changed} परिवर्तन ट्रैक किए गए परिवर्तनों के रूप में मर्ज किए गए',
  reviewCompareIdentical: '{name} के साथ कोई अंतर नहीं: दस्तावेज़ समान हैं',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const hi = {
  reviewCompareMerge: 'तुलना करें (ट्रैक किए गए परिवर्तनों के रूप में मर्ज करें)',
  reviewComparePanel: 'केवल अंतर पैनल दिखाएँ',
  reviewCompareMerged:
    '{name} के साथ तुलना की गई: {added} जोड़, {removed} हटाने और {changed} परिवर्तन ट्रैक किए गए परिवर्तनों के रूप में मर्ज किए गए',
  reviewCompareIdentical: '{name} के साथ कोई अंतर नहीं: दस्तावेज़ समान हैं',
  reviewCompareMergedApprox:
    '{name} के साथ तुलना की गई: {added} जोड़, {removed} हटाने और {changed} परिवर्तन ट्रैक किए गए परिवर्तनों के रूप में मर्ज किए गए (सटीक अनुच्छेद मिलान के लिए दस्तावेज़ बहुत बड़े हैं)',
  reviewCompareDegraded:
    'दस्तावेज़ सटीक अनुच्छेद मिलान के लिए बहुत बड़े हैं: अंतर स्थिति के अनुसार युग्मित किए गए',
} satisfies Record<keyof typeof zh, string>

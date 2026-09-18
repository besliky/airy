import type { zh } from './zh'

export const hi = {
  reviewCompareMerge: 'तुलना करें (ट्रैक किए गए परिवर्तनों के रूप में मर्ज करें)',
  reviewCompareMergeDesc:
    'अंतरों को ट्रैक किए गए परिवर्तनों के रूप में दिखाएँ, जिन्हें स्वीकार या अस्वीकार किया जा सकता है',
  reviewComparePanel: 'केवल अंतर पैनल दिखाएँ',
  reviewComparePanelDesc:
    'दस्तावेज़ में बदलाव किए बिना पैराग्राफ के अंतर साइड पैनल में सूचीबद्ध करता है',
  reviewCompareMerged:
    '{name} के साथ तुलना की गई: {added} जोड़, {removed} हटाने और {changed} परिवर्तन ट्रैक किए गए परिवर्तनों के रूप में मर्ज किए गए',
  reviewCompareIdentical: '{name} के साथ कोई अंतर नहीं: दस्तावेज़ समान हैं',
  reviewCompareMergedApprox:
    '{name} के साथ तुलना की गई: {added} जोड़, {removed} हटाने और {changed} परिवर्तन ट्रैक किए गए परिवर्तनों के रूप में मर्ज किए गए (सटीक अनुच्छेद मिलान के लिए दस्तावेज़ बहुत बड़े हैं)',
  reviewCompareDegraded:
    'दस्तावेज़ सटीक अनुच्छेद मिलान के लिए बहुत बड़े हैं: अंतर स्थिति के अनुसार युग्मित किए गए',
  reviewComparePendingRevisions:
    'दस्तावेज़ में लंबित ट्रैक किए गए परिवर्तन हैं। फिर से तुलना करने से पहले उन्हें स्वीकार या अस्वीकार करें',
  reviewCompareReadonly:
    'तुलना (ट्रैक किए गए परिवर्तनों के रूप में मर्ज करें) के लिए संपादन योग्य दस्तावेज़ चाहिए; यह दस्तावेज़ केवल-पठनीय है',
} satisfies Record<keyof typeof zh, string>

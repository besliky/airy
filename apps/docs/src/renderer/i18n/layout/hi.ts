import type { zh } from './zh'

export const hi = {
  layoutHyphenation: 'हाइफ़नेशन',
  layoutHyphNone: 'कोई नहीं',
  layoutHyphManual: 'मैनुअल',
  layoutHyphAutomatic: 'स्वतः',
  layoutHyphManualDesc: 'जहाँ लाइन तोड़नी है वहाँ वैकल्पिक हाइफ़न डालें',
  layoutHyphManualHint: 'मैनुअल हाइफ़नेशन: वैकल्पिक हाइफ़न डालने के लिए {keys} दबाएँ',
  layoutHyphSet: 'स्वतः हाइफ़नेशन चालू',
  layoutHyphUnset: 'स्वतः हाइफ़नेशन बंद',
  layoutSoftHyphen: 'वैकल्पिक हाइफ़न',
  layoutColsDialogTitle: 'कॉलम',
  layoutColsMore: 'अधिक कॉलम…',
  layoutColOne: 'एक',
  layoutColTwo: 'दो',
  layoutColThree: 'तीन',
  layoutColLeft: 'बाएँ',
  layoutColRight: 'दाएँ',
  layoutColSpacing: 'अंतराल',
  layoutColWidth: 'चौड़ाई',
  layoutColWidth1: 'कॉलम 1 की चौड़ाई',
  layoutLineBetween: 'कॉलम के बीच रेखा',
} satisfies Record<keyof typeof zh, string>

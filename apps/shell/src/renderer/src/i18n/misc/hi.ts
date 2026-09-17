import type { zh } from './zh'

export const hi = {
  today: 'आज',
  yesterday: 'कल',
  daysAgo: '{n} दिन पहले',
  closeTab: 'टैब बंद करें',
  errFileNotFound: 'फ़ाइल मौजूद नहीं है या ले जाई गई है।',
  errPermissionDenied: 'इस फ़ाइल तक पहुँच अस्वीकृत है।',
  errFileLocked: 'फ़ाइल किसी अन्य प्रोग्राम में खुली है। बंद करके पुनः प्रयास करें।',
  errTooManyFiles: 'बहुत अधिक फ़ाइलें खुली हैं। थोड़ी देर बाद पुनः प्रयास करें।',
  tabList: 'सभी टैब',
  newTab: 'नया टैब',
} satisfies Record<keyof typeof zh, string>

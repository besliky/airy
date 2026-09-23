import type { zh } from './zh'

export const he = {
  today: 'היום',
  yesterday: 'אתמול',
  closeTab: 'סגירת כרטיסייה',
  errFileNotFound: 'הקובץ אינו קיים או הועבר.',
  errPermissionDenied: 'הגישה לקובץ זה נדחתה.',
  errIsADirectory: 'זו תיקייה ולא קובץ מסמך.',
  errFileLocked: 'הקובץ פתוח בתוכנית אחרת. סגרו אותה ונסו שוב.',
  errTooManyFiles: 'יותר מדי קבצים פתוחים. נסו שוב בעוד רגע.',
  tabList: 'כל הכרטיסיות',
  newTab: 'כרטיסייה חדשה',
} satisfies Record<keyof typeof zh, string>

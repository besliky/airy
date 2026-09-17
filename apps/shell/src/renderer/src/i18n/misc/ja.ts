import type { zh } from './zh'

export const ja = {
  today: '今日',
  yesterday: '昨日',
  daysAgo: '{n}日前',
  closeTab: 'タブを閉じる',
  errFileNotFound: 'ファイルが存在しないか、移動されました。',
  errPermissionDenied: 'このファイルへのアクセスが拒否されました。',
  errFileLocked: 'ファイルが他のプログラムで開いています。閉じてから再試行してください。',
  errTooManyFiles: '開いているファイルが多すぎます。しばらくしてから再試行してください。',
  tabList: 'すべてのタブ',
  newTab: '新しいタブ',
} satisfies Record<keyof typeof zh, string>

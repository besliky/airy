import type { zh } from './zh'

export const zhTW = {
  today: '今天',
  yesterday: '昨天',
  closeTab: '關閉分頁',
  errFileNotFound: '檔案不存在或已被移動。',
  errPermissionDenied: '沒有存取此檔案的權限。',
  errIsADirectory: '這是資料夾，不是文件檔案。',
  errFileLocked: '檔案正被其他程式佔用，請關閉後再試一次。',
  errTooManyFiles: '開啟的檔案過多，請稍後再試一次。',
  tabList: '全部分頁',
  newTab: '新分頁',
} satisfies Record<keyof typeof zh, string>

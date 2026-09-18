import type { zh } from './zh'

export const zhTW = {
  refsTocMenuAuto: '自動目錄',
  refsTocMenuCustom: '自訂目錄…',
  refsTocMenuTof: '插入圖表目錄…',
  refsTocOptionsTitle: '目錄選項',
  refsTocLevels: '顯示層級',
  refsTocPageNumbers: '顯示頁碼',
  refsTocHyperlinks: '使用超連結',
  refsTocStyles: '從樣式建立',
  refsTocStylesPh: '標題 1,1,標題 2,2',
  refsTofTitle: '圖表目錄',
  refsTofLabel: '標籤',
  refsTofNoCaptions: '文件中沒有此標籤的題目。',
  refsTofFieldLabel: '圖表目錄欄位',
} satisfies Record<keyof typeof zh, string>

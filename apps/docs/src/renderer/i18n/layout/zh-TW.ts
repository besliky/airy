import type { zh } from './zh'

export const zhTW = {
  layoutHyphenation: '斷字',
  layoutHyphNone: '無',
  layoutHyphManual: '手動',
  layoutHyphAutomatic: '自動',
  layoutHyphManualDesc: '在需要斷字的位置插入選擇性連字號',
  layoutHyphManualHint: '手動斷字：按 {keys} 插入選擇性連字號',
  layoutHyphSet: '已開啟自動斷字',
  layoutHyphUnset: '已關閉自動斷字',
  layoutSoftHyphen: '選擇性連字號',
  layoutColsDialogTitle: '欄',
  layoutColsMore: '其他欄…',
  layoutColOne: '一',
  layoutColTwo: '二',
  layoutColThree: '三',
  layoutColLeft: '左',
  layoutColRight: '右',
  layoutColSpacing: '間距',
  layoutColWidth: '欄寬',
  layoutColWidth1: '第 1 欄寬度',
  layoutLineBetween: '欄間分隔線',
} satisfies Record<keyof typeof zh, string>

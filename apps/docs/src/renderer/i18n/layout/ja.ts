import type { zh } from './zh'

export const ja = {
  layoutHyphenation: 'ハイフネーション',
  layoutHyphNone: 'なし',
  layoutHyphManual: '手動',
  layoutHyphAutomatic: '自動',
  layoutHyphManualDesc: '区切りたい位置に任意ハイフンを挿入します',
  layoutHyphManualHint: '手動ハイフネーション：{keys} で任意ハイフンを挿入します',
  layoutHyphSet: '自動ハイフネーションをオンにしました',
  layoutHyphUnset: '自動ハイфネーションをオフにしました',
  layoutSoftHyphen: '任意ハイフン',
  layoutColsDialogTitle: '段組',
  layoutColsMore: '段組の詳細…',
  layoutColOne: '1 段',
  layoutColTwo: '2 段',
  layoutColThree: '3 段',
  layoutColLeft: '左',
  layoutColRight: '右',
  layoutColSpacing: '間隔',
  layoutColWidth: '幅',
  layoutColWidth1: '段 1 の幅',
  layoutLineBetween: '段間罫線',
} satisfies Record<keyof typeof zh, string>

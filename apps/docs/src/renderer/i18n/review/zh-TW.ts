import type { zh } from './zh'

export const zhTW = {
  reviewCompareMerge: '比較（合併為追蹤修訂）',
  reviewComparePanel: '僅顯示差異面板',
  reviewCompareMerged:
    '已與 {name} 比較：{added} 處新增、{removed} 處刪除、{changed} 處修改已合併為追蹤修訂',
  reviewCompareIdentical: '與 {name} 無差異：兩份文件內容相同',
} satisfies Record<keyof typeof zh, string>

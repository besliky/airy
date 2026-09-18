import type { zh } from './zh'

export const zhTW = {
  reviewCompareMerge: '比較（合併為追蹤修訂）',
  reviewComparePanel: '僅顯示差異面板',
  reviewCompareMerged:
    '已與 {name} 比較：{added} 處新增、{removed} 處刪除、{changed} 處修改已合併為追蹤修訂',
  reviewCompareIdentical: '與 {name} 無差異：兩份文件內容相同',
  reviewCompareMergedApprox:
    '已與 {name} 比較：{added} 處新增、{removed} 處刪除、{changed} 處修改已合併為追蹤修訂（文件過大，未做精確段落對齊）',
  reviewCompareDegraded: '文件過大，無法精確對齊段落：差異依位置配對',
  reviewComparePendingRevisions: '文件中還有未處理的追蹤修訂。請先接受或拒絕這些修訂，再進行比較',
  reviewCompareReadonly: '比較（合併為追蹤修訂）需要可編輯的文件；目前文件為唯讀',
} satisfies Record<keyof typeof zh, string>

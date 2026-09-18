import type { zh } from './zh'

export const ja = {
  reviewCompareMerge: '比較（変更履歴にマージ）',
  reviewComparePanel: '差分パネルのみ表示',
  reviewCompareMerged:
    '{name} と比較しました：挿入 {added} 件、削除 {removed} 件、変更 {changed} 件を変更履歴としてマージしました',
  reviewCompareIdentical: '{name} との差分はありません。ドキュメントは同一です',
  reviewCompareMergedApprox:
    '{name} と比較しました：挿入 {added} 件、削除 {removed} 件、変更 {changed} 件を変更履歴としてマージしました（ドキュメントが大きすぎて段落を正確に照合できません）',
  reviewCompareDegraded:
    'ドキュメントが大きすぎて段落を正確に照合できません：差分は位置でペアリングされました',
} satisfies Record<keyof typeof zh, string>

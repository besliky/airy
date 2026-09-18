import type { zh } from './zh'

export const ko = {
  reviewCompareMerge: '비교(수정 사항으로 병합)',
  reviewComparePanel: '차이 패널만 표시',
  reviewCompareMerged:
    '{name} 문서와 비교하여 삽입 {added}건, 삭제 {removed}건, 변경 {changed}건을 수정 사항으로 병합했습니다',
  reviewCompareIdentical: '{name} 문서와 차이가 없습니다. 두 문서가 동일합니다',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const ko = {
  reviewCompareMerge: '비교(수정 사항으로 병합)',
  reviewComparePanel: '차이 패널만 표시',
  reviewCompareMerged:
    '{name} 문서와 비교하여 삽입 {added}건, 삭제 {removed}건, 변경 {changed}건을 수정 사항으로 병합했습니다',
  reviewCompareIdentical: '{name} 문서와 차이가 없습니다. 두 문서가 동일합니다',
  reviewCompareMergedApprox:
    '{name} 문서와 비교하여 삽입 {added}건, 삭제 {removed}건, 변경 {changed}건을 수정 사항으로 병합했습니다(문서가 너무 커서 정확한 단락 일치 불가)',
  reviewCompareDegraded:
    '문서가 너무 커서 단락을 정확하게 일치시킬 수 없습니다: 차이는 위치별로 짝지었습니다',
  reviewComparePendingRevisions:
    '문서에 처리되지 않은 변경 내용이 있습니다. 다시 비교하기 전에 수락하거나 거부하세요',
} satisfies Record<keyof typeof zh, string>

import type { zh } from './zh'

export const ko = {
  reviewCompareMerge: '비교(수정 사항으로 병합)',
  reviewCompareMergeDesc: '차이를 수락하거나 거부할 수 있는 변경 내용으로 표시합니다',
  reviewComparePanel: '차이 패널만 표시',
  reviewComparePanelDesc: '단락 차이를 측면 창에 표시하고 문서는 변경하지 않습니다',
  reviewCompareMerged:
    '{name} 문서와 비교하여 삽입 {added}건, 삭제 {removed}건, 변경 {changed}건을 수정 사항으로 병합했습니다',
  reviewCompareIdentical: '{name} 문서와 차이가 없습니다. 두 문서가 동일합니다',
  reviewCompareMergedApprox:
    '{name} 문서와 비교하여 삽입 {added}건, 삭제 {removed}건, 변경 {changed}건을 수정 사항으로 병합했습니다(문서가 너무 커서 정확한 단락 일치 불가)',
  reviewCompareDegraded:
    '문서가 너무 커서 단락을 정확하게 일치시킬 수 없습니다: 차이는 위치별로 짝지었습니다',
  reviewComparePendingRevisions:
    '문서에 처리되지 않은 변경 내용이 있습니다. 다시 비교하기 전에 수락하거나 거부하세요',
  reviewCompareReadonly:
    '비교(수정 사항으로 병합)에는 편집 가능한 문서가 필요합니다. 이 문서는 읽기 전용입니다',
  reviewComparing: '비교 중…',
} satisfies Record<keyof typeof zh, string>

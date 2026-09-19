import type { zh } from './zh'

export const ko = {
  layoutHyphenation: '하이픈 넣기',
  layoutHyphNone: '없음',
  layoutHyphManual: '수동',
  layoutHyphAutomatic: '자동',
  layoutHyphManualDesc: '끊기를 원하는 위치에 선택적 하이픈을 삽입합니다',
  layoutHyphManualHint: '수동 하이픈 넣기: {keys} 키로 선택적 하이픈을 삽입합니다',
  layoutHyphSet: '자동 하이픈 넣기 켜짐',
  layoutHyphUnset: '자동 하이픈 넣기 꺼짐',
  layoutSoftHyphen: '선택적 하이픈',
  layoutColsDialogTitle: '단',
  layoutColsMore: '단 설정 자세히…',
  layoutColOne: '1단',
  layoutColTwo: '2단',
  layoutColThree: '3단',
  layoutColLeft: '왼쪽',
  layoutColRight: '오른쪽',
  layoutColSpacing: '간격 (cm)',
  layoutColWidth: '너비 (cm)',
  layoutColWidth1: '단 1 너비 (cm)',
  layoutLineBetween: '단 구분선',
} satisfies Record<keyof typeof zh, string>

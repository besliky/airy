import type { zh } from './zh'

export const ko = {
  today: '오늘',
  yesterday: '어제',
  closeTab: '탭 닫기',
  errFileNotFound: '파일이 없거나 이동되었습니다.',
  errPermissionDenied: '이 파일에 대한 액세스 권한이 없습니다.',
  errFileLocked: '파일이 다른 프로그램에서 열려 있습니다. 닫고 다시 시도하세요.',
  errTooManyFiles: '열려 있는 파일이 너무 많습니다. 잠시 후 다시 시도하세요.',
  tabList: '모든 탭',
  newTab: '새 탭',
} satisfies Record<keyof typeof zh, string>

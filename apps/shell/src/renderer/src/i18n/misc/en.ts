import type { zh } from './zh'

export const en = {
  today: 'Today',
  yesterday: 'Yesterday',
  closeTab: 'Close tab',
  errFileNotFound: 'The file does not exist or has been moved.',
  errPermissionDenied: 'Permission denied for this file.',
  errIsADirectory: 'This is a folder, not a document file.',
  errFileLocked: 'The file is open in another program. Close it and try again.',
  errTooManyFiles: 'Too many files are open. Try again shortly.',
  tabList: 'All tabs',
  newTab: 'New tab',
} satisfies Record<keyof typeof zh, string>

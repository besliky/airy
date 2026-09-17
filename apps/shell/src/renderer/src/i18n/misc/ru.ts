import type { zh } from './zh'

export const ru = {
  today: 'Сегодня',
  yesterday: 'Вчера',
  closeTab: 'Закрыть вкладку',
  errFileNotFound: 'Файл не существует или был перемещён.',
  errPermissionDenied: 'Доступ к файлу запрещён.',
  errFileLocked: 'Файл открыт в другой программе. Закройте её и повторите попытку.',
  errTooManyFiles: 'Открыто слишком много файлов. Повторите попытку позже.',
  tabList: 'Все вкладки',
  newTab: 'Новая вкладка',
} satisfies Record<keyof typeof zh, string>

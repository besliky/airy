import type { zh } from './zh'

export const pt = {
  today: 'Hoje',
  yesterday: 'Ontem',
  closeTab: 'Fechar guia',
  errFileNotFound: 'O arquivo não existe ou foi movido.',
  errPermissionDenied: 'Acesso negado a este arquivo.',
  errFileLocked: 'O arquivo está aberto em outro programa. Feche-o e tente novamente.',
  errTooManyFiles: 'Há muitos arquivos abertos. Tente novamente em instantes.',
  tabList: 'Todas as guias',
  newTab: 'Nova guia',
} satisfies Record<keyof typeof zh, string>

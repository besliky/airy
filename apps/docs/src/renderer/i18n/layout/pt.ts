import type { zh } from './zh'

export const pt = {
  layoutHyphenation: 'Hifenização',
  layoutHyphNone: 'Nenhuma',
  layoutHyphManual: 'Manual',
  layoutHyphAutomatic: 'Automática',
  layoutHyphManualDesc: 'Insira hífens opcionais onde quiser quebras',
  layoutHyphManualHint: 'Hifenização manual: pressione {keys} para inserir um hífen opcional',
  layoutHyphSet: 'Hifenização automática ativada',
  layoutHyphUnset: 'Hifenização automática desativada',
  layoutSoftHyphen: 'Hífen opcional',
  layoutColsDialogTitle: 'Colunas',
  layoutColsMore: 'Mais colunas…',
  layoutColOne: 'Uma',
  layoutColTwo: 'Duas',
  layoutColThree: 'Três',
  layoutColLeft: 'Esquerda',
  layoutColRight: 'Direita',
  layoutColSpacing: 'Espaçamento (cm)',
  layoutColWidth: 'Largura (cm)',
  layoutColWidth1: 'Largura da coluna 1 (cm)',
  layoutLineBetween: 'Linha entre colunas',
} satisfies Record<keyof typeof zh, string>

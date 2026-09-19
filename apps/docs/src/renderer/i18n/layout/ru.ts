import type { zh } from './zh'

export const ru = {
  layoutHyphenation: 'Расстановка переносов',
  layoutHyphNone: 'Нет',
  layoutHyphManual: 'Вручную',
  layoutHyphAutomatic: 'Автоматически',
  layoutHyphManualDesc: 'Вставляйте мягкие переносы там, где нужны разрывы',
  layoutHyphManualHint: 'Перенос вручную: нажмите {keys}, чтобы вставить мягкий перенос',
  layoutHyphSet: 'Автоматические переносы включены',
  layoutHyphUnset: 'Автоматические переносы выключены',
  layoutSoftHyphen: 'Мягкий перенос',
  layoutColsDialogTitle: 'Колонки',
  layoutColsMore: 'Другие колонки…',
  layoutColOne: 'Одна',
  layoutColTwo: 'Две',
  layoutColThree: 'Три',
  layoutColLeft: 'Слева',
  layoutColRight: 'Справа',
  layoutColSpacing: 'Интервал (см)',
  layoutColWidth: 'Ширина (см)',
  layoutColWidth1: 'Ширина колонки 1 (см)',
  layoutLineBetween: 'Разделитель колонок',
} satisfies Record<keyof typeof zh, string>

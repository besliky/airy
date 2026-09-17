import type { en } from './en'

/** Fill Form strings (ru); missing keys fall back to the en base. */
export const ru = {
  ribbonTabFillForm: 'Заполнить форму',
  formPreviousField: 'Предыдущее поле',
  formNextField: 'Следующее поле',
  formFieldProgress: '{current} / {total}',
  insertText: 'Вставить текст',
  insertTextHint: 'Вставить в PDF текст, по которому можно искать',
  insertTextTitle: 'Вставить текст',
  editInsertedText: 'Изменить вставленный текст',
  deleteInsertedText: 'Удалить вставленный текст',
  insertedTextDeleted: 'Вставленный текст удалён',
  textInsertSkipped: 'Не удалось сохранить вставленный текст на страницах: {pages}',
  textInsertNoFont:
    'Ни один установленный шрифт не может отрисовать этот текст в PDF (эмодзи и специальные символы не поддерживаются)',
  formComplete: 'Завершить заполнение',
  formMissingRequired: 'Осталось незаполненных обязательных полей: {count}',
  formCompleteDone: 'Проверка заполнения формы пройдена',
  formSignField: 'Нажмите, чтобы подписать',
  formAddText: 'Добавить текст',
  formAddTextHint: 'Введите текст, затем щёлкните по странице, чтобы разместить его',
  formAddTextTitle: 'Добавить текст в PDF',
  formEditText: 'Изменить текст',
  formAddTextPlaceholder: 'Введите текст для размещения',
  formTextSize: 'Размер шрифта',
  formTextColor: 'Цвет',
  formTextAlign: 'Выравнивание',
  formAlignLeft: 'По левому краю',
  formAlignCenter: 'По центру',
  formAlignRight: 'По правому краю',
  formAddCheck: 'Галочка',
  formAddCheckHint: 'Щёлкните по странице, чтобы поставить галочку',
  formAddCross: 'Крестик',
  formAddCrossHint: 'Щёлкните по странице, чтобы поставить крестик',
  formPlaceStaticHint:
    'Щёлкните, чтобы разместить; выберите результат, чтобы переместить или изменить размер',
  formXfaWarning:
    'Этот PDF содержит XFA. Поддерживается только AcroForm; при сохранении данные XFA могут не сохраниться.',
} as const satisfies Partial<Record<keyof typeof en, string>>

import type { en } from './en'

/** Fill Form strings (pl); missing keys fall back to the en base. */
export const pl = {
  ribbonTabFillForm: 'Wypełnij formularz',
  formPreviousField: 'Poprzednie pole',
  formNextField: 'Następne pole',
  formFieldProgress: '{current} / {total}',
  insertText: 'Wstaw tekst',
  insertTextHint: 'Wstaw do pliku PDF tekst, który można przeszukiwać',
  insertTextTitle: 'Wstaw tekst',
  editInsertedText: 'Edytuj wstawiony tekst',
  deleteInsertedText: 'Usuń wstawiony tekst',
  insertedTextDeleted: 'Wstawiony tekst usunięto',
  textInsertSkipped: 'Nie udało się zapisać wstawionego tekstu na stronach: {pages}',
  textInsertNoFont:
    'Żadna zainstalowana czcionka nie może narysować tego tekstu w pliku PDF (emoji i symbole specjalne są nieobsługiwane)',
  formComplete: 'Zakończ wypełnianie',
  formMissingRequired: 'Pozostały puste pola wymagane: {count}',
  formCompleteDone: 'Sprawdzenie wypełnienia formularza zakończyło się powodzeniem',
  formSignField: 'Kliknij, aby podpisać',
  formAddText: 'Dodaj tekst',
  formAddTextHint: 'Wpisz tekst, a następnie kliknij stronę, aby go umieścić',
  formAddTextTitle: 'Dodaj tekst do pliku PDF',
  formEditText: 'Edytuj tekst',
  formAddTextPlaceholder: 'Wpisz tekst do umieszczenia',
  formTextSize: 'Rozmiar czcionki',
  formTextColor: 'Kolor',
  formTextAlign: 'Wyrównanie',
  formAlignLeft: 'Do lewej',
  formAlignCenter: 'Do środka',
  formAlignRight: 'Do prawej',
  formAddCheck: 'Znacznik wyboru',
  formAddCheckHint: 'Kliknij stronę, aby umieścić znacznik wyboru',
  formAddCross: 'Krzyżyk',
  formAddCrossHint: 'Kliknij stronę, aby umieścić znak X',
  formPlaceStaticHint:
    'Kliknij, aby umieścić; zaznacz wynik, aby go przenieść lub zmienić jego rozmiar',
  formXfaWarning:
    'Ten plik PDF zawiera XFA. Obsługiwany jest tylko AcroForm; zapis może nie zachować danych XFA.',
} as const satisfies Partial<Record<keyof typeof en, string>>

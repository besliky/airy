import type { en } from './en'

/** Fill Form strings (cs); missing keys fall back to the en base. */
export const cs = {
  ribbonTabFillForm: 'Vyplnit formulář',
  formPreviousField: 'Předchozí pole',
  formNextField: 'Další pole',
  formFieldProgress: '{current} / {total}',
  insertText: 'Vložit text',
  insertTextHint: 'Vložit do PDF prohledávatelný text',
  insertTextTitle: 'Vložit text',
  editInsertedText: 'Upravit vložený text',
  deleteInsertedText: 'Odstranit vložený text',
  insertedTextDeleted: 'Vložený text byl odstraněn',
  textInsertSkipped: 'Vložený text na stránkách {pages} se nepodařilo uložit',
  textInsertNoFont:
    'Žádné nainstalované písmo nedokáže tento text v PDF vykreslit (emoji a speciální symboly nejsou podporovány)',
  formComplete: 'Dokončit vyplňování',
  formMissingRequired: 'Počet nevyplněných povinných polí: {count}',
  formCompleteDone: 'Kontrola vyplnění formuláře proběhla úspěšně',
  formSignField: 'Kliknutím podepište',
  formAddText: 'Přidat text',
  formAddTextHint: 'Zadejte text a poté klikněte na stránku, kam ho chcete umístit',
  formAddTextTitle: 'Přidat text do PDF',
  formEditText: 'Upravit text',
  formAddTextPlaceholder: 'Zadejte text k umístění',
  formTextSize: 'Velikost písma',
  formTextColor: 'Barva',
  formTextAlign: 'Zarovnání',
  formAlignLeft: 'Vlevo',
  formAlignCenter: 'Na střed',
  formAlignRight: 'Vpravo',
  formAddCheck: 'Zaškrtnutí',
  formAddCheckHint: 'Kliknutím na stránku umístíte zaškrtnutí',
  formAddCross: 'Křížek',
  formAddCrossHint: 'Kliknutím na stránku umístíte křížek',
  formPlaceStaticHint:
    'Kliknutím umístíte; výběrem výsledku ho můžete přesunout nebo změnit jeho velikost',
  formXfaWarning:
    'Tento PDF obsahuje XFA. Podporován je pouze AcroForm; uložení může data XFA nezachovat.',
} as const satisfies Partial<Record<keyof typeof en, string>>

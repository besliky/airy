import type { en } from './en'

/** Fill Form strings (it); missing keys fall back to the en base. */
export const it = {
  ribbonTabFillForm: 'Compila modulo',
  formPreviousField: 'Campo precedente',
  formNextField: 'Campo successivo',
  formFieldProgress: '{current} / {total}',
  insertText: 'Inserisci testo',
  insertTextHint: 'Inserisci nel PDF testo ricercabile',
  insertTextTitle: 'Inserisci testo',
  editInsertedText: 'Modifica testo inserito',
  deleteInsertedText: 'Elimina testo inserito',
  insertedTextDeleted: 'Testo inserito eliminato',
  textInsertSkipped: 'Impossibile salvare il testo inserito nelle pagine: {pages}',
  textInsertNoFont:
    'Nessun carattere installato può disegnare questo testo nel PDF (emoji e simboli speciali non supportati)',
  formComplete: 'Termina compilazione',
  formMissingRequired: '{count} campi obbligatori sono ancora vuoti',
  formCompleteDone: 'Verifica del modulo superata',
  formSignField: 'Fai clic per firmare',
  formAddText: 'Aggiungi testo',
  formAddTextHint: 'Digita il testo, quindi fai clic sulla pagina per posizionarlo',
  formAddTextTitle: 'Aggiungi testo al PDF',
  formEditText: 'Modifica testo',
  formAddTextPlaceholder: 'Immetti il testo da posizionare',
  formTextSize: 'Dimensione carattere',
  formTextColor: 'Colore',
  formTextAlign: 'Allineamento',
  formAlignLeft: 'Sinistra',
  formAlignCenter: 'Centro',
  formAlignRight: 'Destra',
  formAddCheck: 'Segno di spunta',
  formAddCheckHint: 'Fai clic sulla pagina per posizionare un segno di spunta',
  formAddCross: 'Croce',
  formAddCrossHint: 'Fai clic sulla pagina per posizionare una X',
  formPlaceStaticHint:
    'Fai clic per posizionare; seleziona il risultato per spostarlo o ridimensionarlo',
  formXfaWarning:
    'Questo PDF contiene XFA. È supportato solo AcroForm; il salvataggio potrebbe non preservare i dati XFA.',
} as const satisfies Partial<Record<keyof typeof en, string>>

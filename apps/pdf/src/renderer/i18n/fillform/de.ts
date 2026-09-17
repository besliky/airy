import type { en } from './en'

/** Fill Form strings (de); missing keys fall back to the en base. */
export const de = {
  ribbonTabFillForm: 'Formular ausfüllen',
  formPreviousField: 'Vorheriges Feld',
  formNextField: 'Nächstes Feld',
  formFieldProgress: '{current} / {total}',
  insertText: 'Text einfügen',
  insertTextHint: 'Durchsuchbaren Text in das PDF einfügen',
  insertTextTitle: 'Text einfügen',
  editInsertedText: 'Eingefügten Text bearbeiten',
  deleteInsertedText: 'Eingefügten Text löschen',
  insertedTextDeleted: 'Eingefügter Text gelöscht',
  textInsertSkipped: 'Eingefügter Text konnte auf Seite(n) nicht gespeichert werden: {pages}',
  textInsertNoFont:
    'Keine installierte Schriftart kann diesen Text im PDF darstellen (Emoji und Sonderzeichen werden nicht unterstützt)',
  formComplete: 'Ausfüllen beenden',
  formMissingRequired: '{count} Pflichtfelder sind noch leer',
  formCompleteDone: 'Prüfung des Formulars erfolgreich',
  formSignField: 'Zum Signieren klicken',
  formAddText: 'Text hinzufügen',
  formAddTextHint: 'Text eingeben und dann auf die Seite klicken, um ihn zu platzieren',
  formAddTextTitle: 'Text zum PDF hinzufügen',
  formEditText: 'Text bearbeiten',
  formAddTextPlaceholder: 'Einzufügenden Text eingeben',
  formTextSize: 'Schriftgrad',
  formTextColor: 'Farbe',
  formTextAlign: 'Ausrichtung',
  formAlignLeft: 'Links',
  formAlignCenter: 'Zentriert',
  formAlignRight: 'Rechts',
  formAddCheck: 'Häkchen',
  formAddCheckHint: 'Auf die Seite klicken, um ein Häkchen zu platzieren',
  formAddCross: 'Kreuz',
  formAddCrossHint: 'Auf die Seite klicken, um ein X zu platzieren',
  formPlaceStaticHint:
    'Zum Platzieren klicken; das Ergebnis auswählen, um es zu verschieben oder die Größe zu ändern',
  formXfaWarning:
    'Dieses PDF enthält XFA. Nur AcroForm wird unterstützt; beim Speichern gehen XFA-Daten möglicherweise verloren.',
} as const satisfies Partial<Record<keyof typeof en, string>>

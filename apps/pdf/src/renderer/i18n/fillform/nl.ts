import type { en } from './en'

/** Fill Form strings (nl); missing keys fall back to the en base. */
export const nl = {
  ribbonTabFillForm: 'Formulier invullen',
  formPreviousField: 'Vorig veld',
  formNextField: 'Volgend veld',
  formFieldProgress: '{current} / {total}',
  insertText: 'Tekst invoegen',
  insertTextHint: 'Doorzoekbare tekst in de pdf invoegen',
  insertTextTitle: 'Tekst invoegen',
  editInsertedText: 'Ingevoegde tekst bewerken',
  deleteInsertedText: 'Ingevoegde tekst verwijderen',
  insertedTextDeleted: 'Ingevoegde tekst verwijderd',
  textInsertSkipped: "Ingevoegde tekst kon niet worden opgeslagen op pagina('s): {pages}",
  textInsertNoFont:
    'Geen geïnstalleerd lettertype kan deze tekst in de pdf tekenen (emoji en speciale symbolen worden niet ondersteund)',
  formComplete: 'Invullen voltooien',
  formMissingRequired: '{count} verplichte velden zijn nog leeg',
  formCompleteDone: 'Controle van het formulier geslaagd',
  formSignField: 'Klik om te ondertekenen',
  formAddText: 'Tekst toevoegen',
  formAddTextHint: 'Typ tekst en klik vervolgens op de pagina om deze te plaatsen',
  formAddTextTitle: 'Tekst aan pdf toevoegen',
  formEditText: 'Tekst bewerken',
  formAddTextPlaceholder: 'Voer de te plaatsen tekst in',
  formTextSize: 'Lettertypegrootte',
  formTextColor: 'Kleur',
  formTextAlign: 'Uitlijning',
  formAlignLeft: 'Links',
  formAlignCenter: 'Midden',
  formAlignRight: 'Rechts',
  formAddCheck: 'Vinkje',
  formAddCheckHint: 'Klik op de pagina om een vinkje te plaatsen',
  formAddCross: 'Kruis',
  formAddCrossHint: 'Klik op de pagina om een X te plaatsen',
  formPlaceStaticHint:
    'Klik om te plaatsen; selecteer het resultaat om het te verplaatsen of van formaat te veranderen',
  formXfaWarning:
    'Deze pdf bevat XFA. Alleen AcroForm wordt ondersteund; bij het opslaan gaan XFA-gegevens mogelijk verloren.',
} as const satisfies Partial<Record<keyof typeof en, string>>

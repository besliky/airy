import type { en } from './en'

/** Fill Form strings (fr); missing keys fall back to the en base. */
export const fr = {
  ribbonTabFillForm: 'Remplir le formulaire',
  formPreviousField: 'Champ précédent',
  formNextField: 'Champ suivant',
  formFieldProgress: '{current} / {total}',
  insertText: 'Insérer du texte',
  insertTextHint: "Insérer du texte pouvant faire l'objet d'une recherche dans le PDF",
  insertTextTitle: 'Insérer du texte',
  editInsertedText: 'Modifier le texte inséré',
  deleteInsertedText: 'Supprimer le texte inséré',
  insertedTextDeleted: 'Texte inséré supprimé',
  textInsertSkipped: "Le texte inséré n'a pas pu être enregistré sur la ou les pages : {pages}",
  textInsertNoFont:
    'Aucune police installée ne peut dessiner ce texte dans le PDF (les emoji et symboles spéciaux ne sont pas pris en charge)',
  formComplete: 'Terminer le remplissage',
  formMissingRequired: '{count} champs obligatoires sont encore vides',
  formCompleteDone: 'Vérification du formulaire réussie',
  formSignField: 'Cliquer pour signer',
  formAddText: 'Ajouter du texte',
  formAddTextHint: 'Saisissez le texte, puis cliquez sur la page pour le placer',
  formAddTextTitle: 'Ajouter du texte au PDF',
  formEditText: 'Modifier le texte',
  formAddTextPlaceholder: 'Saisir le texte à placer',
  formTextSize: 'Taille de police',
  formTextColor: 'Couleur',
  formTextAlign: 'Alignement',
  formAlignLeft: 'Gauche',
  formAlignCenter: 'Centré',
  formAlignRight: 'Droite',
  formAddCheck: 'Coche',
  formAddCheckHint: 'Cliquer sur la page pour placer une coche',
  formAddCross: 'Croix',
  formAddCrossHint: 'Cliquer sur la page pour placer une croix',
  formPlaceStaticHint:
    'Cliquer pour placer ; sélectionnez le résultat pour le déplacer ou le redimensionner',
  formXfaWarning:
    "Ce PDF contient du XFA. Seul AcroForm est pris en charge ; l'enregistrement peut ne pas préserver les données XFA.",
} as const satisfies Partial<Record<keyof typeof en, string>>

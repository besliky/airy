import type { zh } from './zh'

export const fr = {
  layoutHyphenation: 'Césure',
  layoutHyphNone: 'Aucune',
  layoutHyphManual: 'Manuelle',
  layoutHyphAutomatic: 'Automatique',
  layoutHyphManualDesc: 'Insérez des traits d’union conditionnels aux endroits voulus',
  layoutHyphManualHint:
    'Césure manuelle : appuyez sur {keys} pour insérer un trait d’union conditionnel',
  layoutHyphSet: 'Césure automatique activée',
  layoutHyphUnset: 'Césure automatique désactivée',
  layoutSoftHyphen: 'Trait d’union conditionnel',
  layoutColsDialogTitle: 'Colonnes',
  layoutColsMore: 'Autres colonnes…',
  layoutColOne: 'Une',
  layoutColTwo: 'Deux',
  layoutColThree: 'Trois',
  layoutColLeft: 'Gauche',
  layoutColRight: 'Droite',
  layoutColSpacing: 'Espacement',
  layoutColWidth: 'Largeur',
  layoutColWidth1: 'Largeur colonne 1',
  layoutLineBetween: 'Ligne entre les colonnes',
} satisfies Record<keyof typeof zh, string>

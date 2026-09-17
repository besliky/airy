import type { zh } from './zh'

export const fr = {
  today: "Aujourd'hui",
  yesterday: 'Hier',
  closeTab: "Fermer l'onglet",
  errFileNotFound: "Le fichier n'existe pas ou a été déplacé.",
  errPermissionDenied: 'Accès refusé à ce fichier.',
  errFileLocked: 'Le fichier est ouvert dans un autre programme. Fermez-le puis réessayez.',
  errTooManyFiles: 'Trop de fichiers sont ouverts. Réessayez dans un instant.',
  tabList: 'Tous les onglets',
  newTab: 'Nouvel onglet',
} satisfies Record<keyof typeof zh, string>

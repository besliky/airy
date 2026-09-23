import type { zh } from './zh'

export const es = {
  today: 'Hoy',
  yesterday: 'Ayer',
  closeTab: 'Cerrar pestaña',
  errFileNotFound: 'El archivo no existe o se ha movido.',
  errPermissionDenied: 'Permiso denegado para este archivo.',
  errIsADirectory: 'Esta es una carpeta, no un archivo de documento.',
  errFileLocked: 'El archivo está abierto en otro programa. Ciérralo e inténtalo de nuevo.',
  errTooManyFiles: 'Hay demasiados archivos abiertos. Inténtalo de nuevo en un momento.',
  tabList: 'Todas las pestañas',
  newTab: 'Nueva pestaña',
} satisfies Record<keyof typeof zh, string>

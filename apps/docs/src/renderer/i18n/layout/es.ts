import type { zh } from './zh'

export const es = {
  layoutHyphenation: 'Separación en sílabas',
  layoutHyphNone: 'Ninguna',
  layoutHyphManual: 'Manual',
  layoutHyphAutomatic: 'Automática',
  layoutHyphManualDesc: 'Inserte guiones opcionales donde quiera cortar',
  layoutHyphManualHint: 'Separación manual: pulse {keys} para insertar un guion opcional',
  layoutHyphSet: 'Separación automática activada',
  layoutHyphUnset: 'Separación automática desactivada',
  layoutSoftHyphen: 'Guion opcional',
  layoutColsDialogTitle: 'Columnas',
  layoutColsMore: 'Más columnas…',
  layoutColOne: 'Una',
  layoutColTwo: 'Dos',
  layoutColThree: 'Tres',
  layoutColLeft: 'Izquierda',
  layoutColRight: 'Derecha',
  layoutColSpacing: 'Espaciado',
  layoutColWidth: 'Ancho',
  layoutColWidth1: 'Ancho de columna 1',
  layoutLineBetween: 'Línea entre columnas',
} satisfies Record<keyof typeof zh, string>

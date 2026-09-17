import type { en } from './en'

/** Fill Form strings (es); missing keys fall back to the en base. */
export const es = {
  ribbonTabFillForm: 'Rellenar formulario',
  formPreviousField: 'Campo anterior',
  formNextField: 'Campo siguiente',
  formFieldProgress: '{current} / {total}',
  insertText: 'Insertar texto',
  insertTextHint: 'Insertar texto que se pueda buscar en el PDF',
  insertTextTitle: 'Insertar texto',
  editInsertedText: 'Editar texto insertado',
  deleteInsertedText: 'Eliminar texto insertado',
  insertedTextDeleted: 'Texto insertado eliminado',
  textInsertSkipped: 'No se pudo guardar el texto insertado en las páginas: {pages}',
  textInsertNoFont:
    'Ninguna fuente instalada puede dibujar este texto en el PDF (los emoji y símbolos especiales no son compatibles)',
  formComplete: 'Terminar de rellenar',
  formMissingRequired: '{count} campos obligatorios siguen vacíos',
  formCompleteDone: 'Comprobación del formulario superada',
  formSignField: 'Haga clic para firmar',
  formAddText: 'Agregar texto',
  formAddTextHint: 'Escriba el texto y luego haga clic en la página para colocarlo',
  formAddTextTitle: 'Agregar texto al PDF',
  formEditText: 'Editar texto',
  formAddTextPlaceholder: 'Escriba el texto a colocar',
  formTextSize: 'Tamaño de fuente',
  formTextColor: 'Color',
  formTextAlign: 'Alineación',
  formAlignLeft: 'Izquierda',
  formAlignCenter: 'Centro',
  formAlignRight: 'Derecha',
  formAddCheck: 'Marca de verificación',
  formAddCheckHint: 'Haga clic en la página para colocar una marca de verificación',
  formAddCross: 'Cruz',
  formAddCrossHint: 'Haga clic en la página para colocar una X',
  formPlaceStaticHint:
    'Haga clic para colocar; seleccione el resultado para moverlo o cambiar su tamaño',
  formXfaWarning:
    'Este PDF contiene XFA. Solo se admite AcroForm; al guardar, es posible que los datos XFA no se conserven.',
} as const satisfies Partial<Record<keyof typeof en, string>>

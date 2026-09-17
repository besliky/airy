import type { en } from './en'

/** Fill Form strings (pt); missing keys fall back to the en base. */
export const pt = {
  ribbonTabFillForm: 'Preencher formulário',
  formPreviousField: 'Campo anterior',
  formNextField: 'Próximo campo',
  formFieldProgress: '{current} / {total}',
  insertText: 'Inserir texto',
  insertTextHint: 'Inserir texto pesquisável no PDF',
  insertTextTitle: 'Inserir texto',
  editInsertedText: 'Editar texto inserido',
  deleteInsertedText: 'Excluir texto inserido',
  insertedTextDeleted: 'Texto inserido excluído',
  textInsertSkipped: 'Não foi possível salvar o texto inserido na(s) página(s): {pages}',
  textInsertNoFont:
    'Nenhuma fonte instalada consegue desenhar este texto no PDF (emojis e símbolos especiais não são compatíveis)',
  formComplete: 'Concluir preenchimento',
  formMissingRequired: '{count} campos obrigatórios ainda estão vazios',
  formCompleteDone: 'Verificação do formulário aprovada',
  formSignField: 'Clique para assinar',
  formAddText: 'Adicionar texto',
  formAddTextHint: 'Digite o texto e clique na página para posicioná-lo',
  formAddTextTitle: 'Adicionar texto ao PDF',
  formEditText: 'Editar texto',
  formAddTextPlaceholder: 'Digite o texto a posicionar',
  formTextSize: 'Tamanho da fonte',
  formTextColor: 'Cor',
  formTextAlign: 'Alinhamento',
  formAlignLeft: 'Esquerda',
  formAlignCenter: 'Centro',
  formAlignRight: 'Direita',
  formAddCheck: 'Marca de verificação',
  formAddCheckHint: 'Clique na página para posicionar uma marca de verificação',
  formAddCross: 'X',
  formAddCrossHint: 'Clique na página para posicionar um X',
  formPlaceStaticHint:
    'Clique para posicionar; selecione o resultado para movê-lo ou redimensioná-lo',
  formXfaWarning:
    'Este PDF contém XFA. Somente o AcroForm é compatível; salvar pode não preservar os dados do XFA.',
} as const satisfies Partial<Record<keyof typeof en, string>>

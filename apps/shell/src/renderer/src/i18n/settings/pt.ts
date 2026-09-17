import type { zh } from './zh'

export const pt = {
  versionLabel: 'Versão',
  theme: 'Tema',
  themeLight: 'Claro',
  themeDark: 'Escuro',
  themeSystem: 'Seguir o Sistema',
  setAuthorName: 'Nome do autor',
  setAuthorNameDesc: 'Usado em comentários e alterações controladas.',
  setAuthorNamePlaceholder: 'Usuário',
  saveLocation: 'Local de salvamento',
  setAutoSave: 'Salvar automaticamente todos os documentos',
  setAutoSaveDesc:
    'Ativa o salvamento automático por padrão em todos os editores. Você ainda pode desativá-lo em uma janela específica.',
  setRestoreSession: 'Restaurar sessão anterior',
  setRestoreSessionDesc:
    'Reabre ao iniciar os arquivos abertos no último encerramento ou travamento.',
  setLiveBridge: 'Ponte Copilot ao vivo',
  setLiveBridgeDesc:
    'Permite que agentes de código (MCP) se conectem a este aplicativo para ler e editar documentos abertos. Usado por ferramentas como Claude, Cursor ou airy-mcp.',
  setLiveBridgeEnvDisabled:
    'Desativado à força pela variável de ambiente AIRY_DISABLE_BRIDGE=1; este interruptor não tem efeito até que ela seja removida.',
  setAiFontSize: 'Tamanho do texto do painel de IA',
  aiFontSizeDefault: 'Padrão',
  aiFontSizeLarge: 'Grande',
  aiFontSizeXLarge: 'Muito grande',
  aiFontSizeCustom: 'Personalizado',
  setAiSpellcheck: 'Verificação ortográfica no chat de IA',
  setAiSpellcheckDesc:
    'Sublinha palavras com erros ortográficos ao digitar na caixa do chat de IA.',
  settings: 'Configurações',
  setSecGeneral: 'Geral',
  setSecAbout: 'Sobre',
  setSecAiModel: 'Modelo de IA',
  setAiProvider: 'Provedor',
  setAiModelId: 'Modelo',
  setAiApiKey: 'Chave de API',
  setAiKeyHint: 'Armazenada apenas neste dispositivo.',
  setAiBaseUrl: 'Base URL',
  setAiBaseUrlHint: 'Deixe vazio para o endpoint oficial.',
  setAiNoneHint: 'Nenhum provedor configurado. Escolha um provedor e insira a chave de API.',
  setAiCodexPath: 'Executável do Codex',
  setAiCodexPathHint:
    'Preencha apenas para uma instalação personalizada; deixe em branco para detectar automaticamente.',
  setAiCodexAutoPlaceholder: 'Detectar automaticamente (recomendado)',
  setAiCodexHint: 'Usa o Codex CLI conectado localmente; nenhuma chave de API é necessária.',
  setAiByokNote:
    'Os chats usam sua própria chave. A geração de imagens e a análise de mídia seguem a seção "Mídia de IA"; a busca na web usa chaves próprias ou fontes gratuitas.',
  setAiSave: 'Salvar',
  setAiSaved: 'Salvo',
  setAiTest: 'Testar conexão',
  setAiTesting: 'Testando…',
  setAiTestOk: 'Conexão bem-sucedida',
  setAiTestFail: 'Falha na conexão',
  setAiMaxTokens: 'Máx. de tokens de saída',
  setAiMaxTokensDesc:
    'Orçamento de saída por turno. Modelos de raciocínio gastam-no pensando; se esgotar, a resposta vem vazia — aumente este valor.',
  setSecAiMedia: 'Mídia e busca de IA',
  setAiSearchSerperHint: 'O Serper oferece busca na web e de imagens com a sua chave.',
  setAiSearchTavilyHint:
    'O Tavily oferece busca na web com a sua chave; a busca de imagens recorre a fontes gratuitas.',
  setAiCapImage: 'Geração de imagens',
  setAiCapAnalysis: 'Análise de imagens',
  setAiCapVideo: 'Análise de vídeo',
  setAiCapSearch: 'Busca na web',
  setAiSharedKeyHint:
    'A chave e a URL base de um provedor são compartilhadas entre as capacidades; insira-as uma só vez.',
  setGithub: 'Código aberto',
  starPromptTitle: 'Gostando do Airy?',
  starPromptTitleN: 'Você já abriu {n} documentos com o Airy',
  starPromptBody:
    'O Airy é gratuito e de código aberto. Uma estrela no GitHub é a melhor forma de apoiar a equipe.',
  starPromptGo: 'Dar uma estrela',
  starPromptDone: 'Já dei',
  starPromptLater: 'Mais tarde',
  setChange: 'Alterar',
  language: 'Idioma',
} satisfies Record<keyof typeof zh, string>

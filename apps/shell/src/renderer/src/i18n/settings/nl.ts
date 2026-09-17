import type { zh } from './zh'

export const nl = {
  versionLabel: 'Versie',
  theme: 'Thema',
  themeLight: 'Licht',
  themeDark: 'Donker',
  themeSystem: 'Systeem volgen',
  setAuthorName: 'Auteursnaam',
  setAuthorNameDesc: 'Gebruikt voor opmerkingen en wijzigingen bijhouden.',
  setAuthorNamePlaceholder: 'Gebruiker',
  saveLocation: 'Opslaglocatie',
  setAutoSave: 'Alle documenten automatisch opslaan',
  setAutoSaveDesc:
    'Zet automatisch opslaan standaard aan in elke editor. Je kunt het nog steeds uitschakelen voor één venster.',
  setRestoreSession: 'Vorige sessie herstellen',
  setRestoreSessionDesc:
    'Opent bij het opstarten de bestanden die openstonden bij het laatste afsluiten of crashen.',
  setLiveBridge: 'Copilot live bridge',
  setLiveBridgeDesc:
    'Laat coding-agenten (MCP) verbinden met deze app om geopende documenten te lezen en bewerken. Gebruikt door tools zoals Claude, Cursor of airy-mcp.',
  setLiveBridgeEnvDisabled:
    'Uitgedwongen uit door de omgevingsvariabele AIRY_DISABLE_BRIDGE=1; deze schakelaar doet niets tot die wordt verwijderd.',
  setAiFontSize: 'Tekstgrootte AI-paneel',
  aiFontSizeDefault: 'Standaard',
  aiFontSizeLarge: 'Groot',
  aiFontSizeXLarge: 'Extra groot',
  aiFontSizeCustom: 'Aangepast',
  setAiSpellcheck: 'Spellingcontrole in AI-chat',
  setAiSpellcheckDesc:
    'Onderstreept verkeerd gespelde woorden tijdens het typen in het AI-chatveld.',
  settings: 'Instellingen',
  setSecGeneral: 'Algemeen',
  setSecAbout: 'Over',
  setSecAiModel: 'AI-model',
  setAiProvider: 'Provider',
  setAiModelId: 'Model',
  setAiApiKey: 'API-sleutel',
  setAiKeyHint: 'Alleen op dit apparaat opgeslagen.',
  setAiBaseUrl: 'Base URL',
  setAiBaseUrlHint: 'Leeg laten voor het officiële eindpunt.',
  setAiNoneHint: 'Geen provider ingesteld. Kies een provider en voer de API-sleutel in.',
  setAiCodexPath: 'Codex-uitvoerbaar bestand',
  setAiCodexPathHint:
    'Alleen invullen voor een aangepaste installatie; laat leeg voor automatische detectie.',
  setAiCodexAutoPlaceholder: 'Automatisch detecteren (aanbevolen)',
  setAiCodexHint: 'Gebruikt de lokaal aangemelde Codex CLI; geen API-sleutel nodig.',
  setAiByokNote:
    'Chats gebruiken je eigen sleutel. Beeldgeneratie en media-analyse volgen de sectie "AI-media"; webzoekopdrachten gebruiken eigen sleutels of gratis bronnen.',
  setAiSave: 'Opslaan',
  setAiSaved: 'Opgeslagen',
  setAiTest: 'Verbinding testen',
  setAiTesting: 'Testen…',
  setAiTestOk: 'Verbinding geslaagd',
  setAiTestFail: 'Verbinding mislukt',
  setAiMaxTokens: 'Max. outputtokens',
  setAiMaxTokensDesc:
    'Uitvoerbudget voor één beurt. Redeneermodellen geven dit uit aan denken; is het op, dan komt een leeg antwoord terug — verhoog deze waarde.',
  setSecAiMedia: 'AI-media en zoeken',
  setAiSearchSerperHint: 'Serper levert web- en afbeeldingszoeken met je sleutel.',
  setAiSearchTavilyHint:
    'Tavily levert webzoeken met je sleutel; afbeeldingszoeken valt terug op gratis bronnen.',
  setAiCapImage: 'Afbeeldingen genereren',
  setAiCapAnalysis: 'Afbeeldingsanalyse',
  setAiCapVideo: 'Video-analyse',
  setAiCapSearch: 'Zoeken op het web',
  setAiSharedKeyHint:
    'De sleutel en basis-URL van een provider gelden voor alle functies; één keer invoeren volstaat.',
  setGithub: 'Open source',
  starPromptTitle: 'Bevalt Airy?',
  starPromptTitleN: 'Je hebt al {n} documenten geopend met Airy',
  starPromptBody:
    'Airy is gratis en open source. Een ster op GitHub is de beste steun voor het team.',
  starPromptGo: 'Geef een ster',
  starPromptDone: 'Al gedaan',
  starPromptLater: 'Later',
  setChange: 'Wijzigen',
  language: 'Taal',
} satisfies Record<keyof typeof zh, string>

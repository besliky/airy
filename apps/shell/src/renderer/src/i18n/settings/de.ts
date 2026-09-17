import type { zh } from './zh'

export const de = {
  versionLabel: 'Version',
  theme: 'Thema',
  themeLight: 'Hell',
  themeDark: 'Dunkel',
  themeSystem: 'System folgen',
  setAuthorName: 'Autorname',
  setAuthorNameDesc: 'Wird für Kommentare und nachverfolgte Änderungen verwendet.',
  setAuthorNamePlaceholder: 'Benutzer',
  saveLocation: 'Speicherort',
  setAutoSave: 'Alle Dokumente automatisch speichern',
  setAutoSaveDesc:
    'Aktiviert AutoSave standardmäßig in jedem Editor. Für ein einzelnes Fenster kann es weiterhin ausgeschaltet werden.',
  setRestoreSession: 'Vorherige Sitzung wiederherstellen',
  setRestoreSessionDesc:
    'Öffnet beim Start die Dateien erneut, die beim letzten Schließen oder Absturz geöffnet waren.',
  setLiveBridge: 'Copilot-Live-Bridge',
  setLiveBridgeDesc:
    'Erlaubt Coding-Agenten (MCP), sich mit dieser App zu verbinden und geöffnete Dokumente zu lesen und zu bearbeiten. Wird von Tools wie Claude, Cursor oder airy-mcp genutzt.',
  setLiveBridgeEnvDisabled:
    'Durch die Umgebungsvariable AIRY_DISABLE_BRIDGE=1 erzwungen aus; der Schalter bleibt wirkungslos, bis sie entfernt wird.',
  setAiFontSize: 'Textgröße im KI-Bereich',
  aiFontSizeDefault: 'Standard',
  aiFontSizeLarge: 'Groß',
  aiFontSizeXLarge: 'Sehr groß',
  aiFontSizeCustom: 'Benutzerdefiniert',
  setAiSpellcheck: 'Rechtschreibprüfung im KI-Chat',
  setAiSpellcheckDesc:
    'Unterstreicht falsch geschriebene Wörter beim Tippen im KI-Chat-Eingabefeld.',
  settings: 'Einstellungen',
  setSecGeneral: 'Allgemein',
  setSecAbout: 'Über',
  setSecAiModel: 'KI-Modell',
  setAiProvider: 'Anbieter',
  setAiModelId: 'Modell',
  setAiApiKey: 'API-Schlüssel',
  setAiKeyHint: 'Wird nur auf diesem Gerät gespeichert.',
  setAiBaseUrl: 'Base URL',
  setAiBaseUrlHint: 'Leer lassen für den offiziellen Endpunkt.',
  setAiNoneHint:
    'Kein Anbieter konfiguriert. Wählen Sie einen Anbieter und geben Sie dessen API-Schlüssel ein.',
  setAiCodexPath: 'Codex-Programmdatei',
  setAiCodexPathHint:
    'Nur bei einer benutzerdefinierten Installation angeben; leer lassen für automatische Erkennung.',
  setAiCodexAutoPlaceholder: 'Automatisch erkennen (empfohlen)',
  setAiCodexHint: 'Verwendet die lokal angemeldete Codex CLI; kein API-Schlüssel nötig.',
  setAiByokNote:
    'Chats nutzen deinen eigenen Schlüssel. Bildgenerierung und Medienanalyse folgen dem Abschnitt „KI-Medien“; die Websuche nutzt eigene Schlüssel oder kostenlose Quellen.',
  setAiSave: 'Speichern',
  setAiSaved: 'Gespeichert',
  setAiTest: 'Verbindung testen',
  setAiTesting: 'Wird getestet…',
  setAiTestOk: 'Verbindung erfolgreich',
  setAiTestFail: 'Verbindung fehlgeschlagen',
  setAiMaxTokens: 'Max. Ausgabe-Tokens',
  setAiMaxTokensDesc:
    'Ausgabe-Budget pro Durchlauf. Denk-Modelle verbrauchen es beim Reasoning; ist es erschöpft, kommt eine leere Antwort zurück — dann diesen Wert erhöhen.',
  setSecAiMedia: 'KI-Medien & Suche',
  setAiSearchSerperHint: 'Serper liefert mit deinem Schlüssel Web- und Bildsuche.',
  setAiSearchTavilyHint:
    'Tavily liefert mit deinem Schlüssel die Websuche; die Bildsuche greift auf kostenlose Quellen zurück.',
  setAiCapImage: 'Bildgenerierung',
  setAiCapAnalysis: 'Bildanalyse',
  setAiCapVideo: 'Videoanalyse',
  setAiCapSearch: 'Websuche',
  setAiSharedKeyHint:
    'Schlüssel und Base URL eines Anbieters gelten für alle Fähigkeiten; einmal eintragen genügt.',
  setGithub: 'Open Source',
  starPromptTitle: 'Gefällt Ihnen Airy?',
  starPromptTitleN: 'Sie haben {n} Dokumente mit Airy geöffnet',
  starPromptBody:
    'Airy ist kostenlos und Open Source. Ein Stern auf GitHub ist die beste Unterstützung für das Team.',
  starPromptGo: 'Stern geben',
  starPromptDone: 'Schon erledigt',
  starPromptLater: 'Später',
  setChange: 'Ändern',
  language: 'Sprache',
} satisfies Record<keyof typeof zh, string>

import type { AiPanelPrefs } from '@airy-office/ui'
import type { Lang } from '@airy-office/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@airy-office/ai-provider'

export const MARKDOWN_CHANNELS = {
  consumePending: 'markdown:consume-pending',
  readFile: 'markdown:read-file',
  writeRecovery: 'markdown:write-recovery',
  save: 'markdown:save',
  saveRequest: 'markdown:save-request',
  saveRequestAck: 'markdown:save-request-ack',
  setEncoding: 'markdown:set-encoding',
  dirtyChanged: 'markdown:dirty-changed',
  closeSaveRequest: 'markdown:close-save-request',
  closeSaveResult: 'markdown:close-save-result',
  fileRenamed: 'markdown:file-renamed',
  pickImage: 'markdown:pick-image',
  saveImage: 'markdown:save-image',
  readImage: 'markdown:read-image',
  exportRequest: 'markdown:export-request',
  exportDocx: 'markdown:export-docx',
  exportPdf: 'markdown:export-pdf',
  exportHtml: 'markdown:export-html',
  printRequest: 'markdown:print-request',
  aiGenerateImage: 'markdown:ai-generate-image',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
  getAutoSaveDefault: 'app:get-auto-save-default',
  autoSaveDefaultChanged: 'app:auto-save-default-changed',
  getAiPanelPrefs: 'app:get-ai-panel-prefs',
  aiPanelPrefsChanged: 'app:ai-panel-prefs-changed',
} as const

export type UiTheme = 'light' | 'dark' | 'system'

/**
 * Charsets a manual "Reopen with encoding" pick may name (UX-1696). Mirrors
 * the detector's candidate set (LEGACY_CHARSETS in
 * packages/file-parse/src/text.ts) plus the UTF-8/16 family a user may want
 * to force over a wrong legacy guess. One list shared by the renderer (picker
 * options), the preload (pass-through) and main (persistence + validation).
 * Keep in sync with LEGACY_CHARSETS; a persisted entry whose encoding is not
 * listed here is dropped on read and the file falls back to auto-detection.
 */
export const SELECTABLE_ENCODINGS = [
  'utf-8',
  'utf-16le',
  'utf-16be',
  'gb18030',
  'shift_jis',
  'big5',
  'euc-kr',
  'windows-1252',
  'windows-1251',
  'koi8-r',
  'windows-1250',
  'windows-1253',
  'windows-1255',
  'windows-1256',
  'windows-874',
] as const

export type SelectableEncoding = (typeof SELECTABLE_ENCODINGS)[number]

/** shell-wide AutoSave default; updatedAt is 0 until the user has ever set it */
export interface AutoSaveDefault {
  on: boolean
  updatedAt: number
}

export type SaveMode = 'save' | 'saveAs'

export interface SaveMarkdownRequest {
  /** full document text (frontmatter included) */
  text: string
  /** Authored image paths in document order; the main process validates every path. */
  imageSources: string[]
  mode: SaveMode
  /**
   * Silent first save for an untitled document (AI auto-naming): saves to a
   * unique path under Documents derived from this name, without a dialog.
   * Ignored when the document already has a path.
   */
  suggestedName?: string
  /**
   * Marks an automatic save (the autosave tick / window blur). When the
   * staleness fence refuses the write, an auto save is declined without the
   * external-change dialog instead of popping a modal every 30 seconds.
   */
  auto?: boolean
}

export type SaveMarkdownResult =
  | {
      ok: true
      path: string
      /** Save As may relocate local images into the new document's assets directory. */
      imageRewrites?: Array<{ from: string; to: string }>
    }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** AI channels are app-wide shared ipcMain handlers (shell registers via docs-main registerAiIpc); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
  webSearch: 'ai:web-search',
  imageSearch: 'ai:image-search',
  fetchImage: 'ai:fetch-image',
} as const

export interface WebSearchResult {
  answer?: string
  results: Array<{ title: string; url: string; snippet: string }>
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

export interface ImageSearchResult {
  images: Array<{ title?: string; imageUrl: string; width?: number; height?: number }>
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

export type ExportFormat = 'pdf' | 'docx' | 'docs' | 'html'

export interface ExportDocxRequest {
  /** .docx bytes, base64 */
  base64: string
  /** file name (no extension) suggested in the dialog / used for the silent convert */
  suggestedName: string
  /** 'dialog' = save dialog; 'openInDocs' = app-managed temporary copy opened in AI Docs */
  mode: 'dialog' | 'openInDocs'
}

export interface ExportPdfRequest {
  /** self-contained print HTML */
  html: string
  suggestedName: string
}

export interface ExportHtmlRequest {
  /** standalone HTML document (inline CSS, offline-openable) */
  html: string
  suggestedName: string
}

export type ExportResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

export interface ImageData {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
}

/** API exposed by preload to the renderer (window.markdownApi) */
export interface MarkdownApi {
  /** Take the md path pending for this view (queued at tab creation); null = new untitled document */
  consumePending(): Promise<string | null>
  /** Read the file as UTF-8 text. Only paths granted to this view are allowed */
  readFile(path: string): Promise<{ text: string; recovered: boolean }>
  /**
   * Remember (encoding) or forget (null, back to auto-detection) the charset
   * this file decodes with; the caller re-reads via readFile afterwards
   * (UX-1696 "Reopen with encoding"). Only paths granted to this view are
   * allowed; the pick survives tab close and relaunch (UX-1653).
   */
  setEncoding(path: string, encoding: string | null): Promise<boolean>
  /** crash-recovery copy push (dirty renderers, every ~30s and on blur) */
  writeRecovery(path: string, text: string): Promise<void>
  /**
   * Write the document text. With a granted file path the write is atomic
   * (tmp + rename); untitled documents and mode 'saveAs' go through a main-process
   * save dialog first. The resolved path is granted to the view and returned.
   */
  save(request: SaveMarkdownRequest): Promise<SaveMarkdownResult>
  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */
  setDirty(dirty: boolean): void
  /** Shell menu Save / Save As → renderer serializes and calls save() with the given mode */
  onSaveRequest(handler: (mode: SaveMode) => void): () => void
  /** Resolves a menu-save waiter when doSave exits without ever invoking save() (busy/loading) */
  sendSaveRequestAck(ok: boolean): void
  /** Main process picked "Save" in the close prompt → renderer saves and replies via sendCloseSaveResult */
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  /** The file was renamed on disk (Home list rename) — renderer syncs its display path */
  onFileRenamed(handler: (newPath: string) => void): () => void
  /**
   * Pick an image file and copy it into `assets/` next to the open document;
   * returns the relative path to author into the markdown, or null when the
   * document is untitled or the picker was canceled.
   */
  pickImage(): Promise<string | null>
  /**
   * Persist pasted/dropped image bytes into `assets/` next to the open
   * document; returns the relative path to author, or null when untitled.
   */
  saveImage(data: { base64: string; ext: string }): Promise<string | null>
  /**
   * Read an image referenced by the document for DOCX embedding. Only paths
   * inside the document's directory are allowed; anything else returns null.
   */
  readImage(src: string): Promise<ImageData | null>
  /** Shell menu export → renderer serializes and calls exportDocx/exportPdf */
  onExportRequest(handler: (format: ExportFormat) => void): () => void
  /** Shell menu Print → renderer builds the print HTML and opens the system print dialog */
  onPrintRequest(handler: () => void): () => void
  exportDocx(request: ExportDocxRequest): Promise<ExportResult>
  exportPdf(request: ExportPdfRequest): Promise<ExportResult>
  /** Standalone HTML export: save-dialog path, written atomically by main */
  exportHtml(request: ExportHtmlRequest): Promise<ExportResult>
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  getAutoSaveDefault(): Promise<AutoSaveDefault>
  onAutoSaveDefaultChanged(handler: (value: AutoSaveDefault) => void): () => void
  /** AI panel text size + chat-input spellcheck (Settings → General in the shell) */
  getAiPanelPrefs(): Promise<AiPanelPrefs>
  onAiPanelPrefsChanged(handler: (prefs: AiPanelPrefs) => void): () => void
  /** press on the shell chrome (tab strip is a sibling WebContentsView whose
   *  clicks produce no DOM event here) — dismiss open popovers */
  onChromePressed(handler: () => void): () => void
  getAiSettings(): Promise<AiSettings>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /** Main-process web search (Serper/DuckDuckGo via the shared ai:web-search handler) */
  webSearch(query: string, maxResults?: number): Promise<WebSearchResult>
  /** Main-process image search (shared ai:image-search handler) */
  imageSearch(query: string, maxResults?: number): Promise<ImageSearchResult>
  /** Download an image URL in the main process (CORS-free, scheme/target validated) */
  fetchImage(url: string): Promise<{ base64: string; mime: string } | null>
  /** AI image generation via the configured media provider (markdown-owned channel) */
  aiGenerateImage(op: { prompt: string; aspectRatio?: string }): Promise<{
    url?: string
    error?: string
  }>
}
